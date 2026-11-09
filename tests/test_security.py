"""Regression tests for SSRF / arbitrary local file read via <img> and <link>."""
import http.server
import os
import socket
import socketserver
import tempfile
import threading
import unittest
from unittest import mock

from docx import Document

from html4docx import HtmlToDocx, resources

test_dir = os.path.abspath(os.path.dirname(__file__))
TEST_IMG = os.path.join(test_dir, "assets", "images", "test_img.png")


class _Handler(http.server.BaseHTTPRequestHandler):
    hits = []

    def do_GET(self):
        _Handler.hits.append(self.path)
        if self.path.startswith("/redirect"):
            self.send_response(302)
            self.send_header("Location", self.path.split("to=", 1)[1])
            self.end_headers()
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/css")
        self.end_headers()
        self.wfile.write(b"p { color: red; }")

    def log_message(self, *args):
        pass


class SecurityTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = socketserver.TCPServer(("127.0.0.1", 0), _Handler)
        cls.port = cls.server.server_address[1]
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def setUp(self):
        _Handler.hits.clear()
        fd, self.secret = tempfile.mkstemp()
        os.write(fd, b"SECRET_FILE_CONTENTS")
        os.close(fd)

    def tearDown(self):
        os.remove(self.secret)

    def _url(self, path):
        return f"http://127.0.0.1:{self.port}{path}"

    # --- SSRF ---------------------------------------------------------- #
    def test_default_blocks_loopback_ssrf_via_img_and_link(self):
        html = f'<link rel="stylesheet" href="{self._url("/css")}"><img src="{self._url("/img")}">'
        HtmlToDocx().add_html_to_document(html, Document())
        self.assertEqual(_Handler.hits, [])

    def test_default_blocks_metadata_and_private_ips(self):
        for url in (
            "http://169.254.169.254/latest/meta-data/",
            "http://10.0.0.1/",
            "http://192.168.0.1/",
            "http://[::1]/",
            "http://[::ffff:127.0.0.1]/",
            "http://0.0.0.0/",
            "http://localhost/",
        ):
            with self.subTest(url=url), self.assertRaises(resources.BlockedResourceError):
                resources._fetch_remote_bytes(url, resources.ResourcePolicy())

    def test_redirect_to_private_address_is_blocked(self):
        # Simulates a public first hop by allowing only the first connection.
        target = self._url("/secret")
        real = resources._guarded_create_connection
        calls = {"n": 0}

        def first_hop_allowed(address, *args, **kwargs):
            calls["n"] += 1
            if calls["n"] == 1:
                return socket.create_connection(address, *args[:1])
            return real(address, *args, **kwargs)

        with mock.patch.object(resources, "_guarded_create_connection", first_hop_allowed):
            result = resources.fetch_external_css(self._url(f"/redirect?to={target}"))
        self.assertIsNone(result)
        self.assertEqual(_Handler.hits, ["/redirect?to=" + target])

    def test_non_http_schemes_are_blocked(self):
        for src in ("ftp://example.com/a.css", "gopher://127.0.0.1:70/", "jar:file:///x!/y"):
            with self.subTest(src=src):
                self.assertIsNone(resources.fetch_external_css(src))

    def test_opt_in_private_network(self):
        policy = resources.ResourcePolicy(allow_private_network=True)
        self.assertEqual(resources.fetch_external_css(self._url("/css"), policy), "p { color: red; }")

    def test_remote_can_be_disabled(self):
        parser = HtmlToDocx()
        parser.options["allow-remote-resources"] = False
        parser.options["allow-private-network"] = True
        parser.add_html_to_document(f'<img src="{self._url("/img")}">', Document())
        self.assertEqual(_Handler.hits, [])

    def test_max_size_enforced(self):
        policy = resources.ResourcePolicy(allow_private_network=True, max_bytes=5)
        self.assertIsNone(resources.fetch_external_css(self._url("/css"), policy))

    # --- Private network allowlist ----------------------------------- #
    def test_allowlist_allows_listed_private_ip(self):
        policy = resources.ResourcePolicy(private_network_allowlist=["127.0.0.1/32"])
        self.assertEqual(resources.fetch_external_css(self._url("/css"), policy), "p { color: red; }")

    def test_allowlist_via_options(self):
        parser = HtmlToDocx()
        parser.options["private-network-allowlist"] = ["127.0.0.1"]
        parser.add_html_to_document(f'<link rel="stylesheet" href="{self._url("/css")}"><p>x</p>', Document())
        self.assertEqual(_Handler.hits, ["/css"])

    def test_allowlist_does_not_open_other_private_ips(self):
        allowlist = resources._parse_allowlist(["10.0.5.0/24"])
        self.assertFalse(resources._is_forbidden_ip("10.0.5.20", allowlist))
        self.assertTrue(resources._is_forbidden_ip("10.0.6.1", allowlist))
        self.assertTrue(resources._is_forbidden_ip("127.0.0.1", allowlist))
        self.assertFalse(resources._is_forbidden_ip("8.8.8.8", allowlist))  # public still allowed

    def test_allow_private_network_takes_precedence_over_allowlist(self):
        # 127.0.0.1 is NOT in the allowlist, but allow-private-network=True opens everything
        parser = HtmlToDocx()
        parser.options["allow-private-network"] = True
        parser.options["private-network-allowlist"] = ["10.0.5.20"]
        parser.add_html_to_document(f'<link rel="stylesheet" href="{self._url("/css")}"><p>x</p>', Document())
        self.assertEqual(_Handler.hits, ["/css"])

    def test_allowlist_alone_does_not_open_unlisted_private_ip(self):
        parser = HtmlToDocx()
        parser.options["private-network-allowlist"] = ["10.0.5.20"]
        parser.add_html_to_document(f'<link rel="stylesheet" href="{self._url("/css")}"><p>x</p>', Document())
        self.assertEqual(_Handler.hits, [])

    def test_disabling_remote_overrides_all_network_options(self):
        parser = HtmlToDocx()
        parser.options["allow-remote-resources"] = False
        parser.options["allow-private-network"] = True
        parser.options["private-network-allowlist"] = ["127.0.0.1"]
        parser.add_html_to_document(f'<link rel="stylesheet" href="{self._url("/css")}"><p>x</p>', Document())
        self.assertEqual(_Handler.hits, [])

    def test_warns_when_allow_private_network_and_allowlist_are_combined(self):
        with self.assertLogs("html4docx.resources", level="WARNING") as logs:
            resources.ResourcePolicy(allow_private_network=True, private_network_allowlist=["10.0.5.20"])
        self.assertTrue(any("'private-network-allowlist' is ignored" in line for line in logs.output))

    def test_warns_when_combined_through_parser_options(self):
        parser = HtmlToDocx()
        parser.options["allow-private-network"] = True
        parser.options["private-network-allowlist"] = ["10.0.5.20"]
        with self.assertLogs("html4docx.resources", level="WARNING") as logs:
            parser._resource_policy()
        self.assertTrue(any("'private-network-allowlist' is ignored" in line for line in logs.output))

    def test_no_warning_when_only_one_network_option_is_used(self):
        cases = {
            "only allow-private-network": {"allow_private_network": True},
            "only private-network-allowlist": {"private_network_allowlist": ["10.0.5.20"]},
            "neither (defaults)": {},
        }
        for name, kwargs in cases.items():
            with self.subTest(name):
                # assertNoLogs needs Python 3.10+, so patch the logger to stay compatible with 3.9
                with mock.patch.object(resources.logger, "warning") as warning:
                    resources.ResourcePolicy(**kwargs)
                warning.assert_not_called()

    def test_invalid_allowlist_entry_fails_fast(self):
        with self.assertRaises(ValueError):
            resources.ResourcePolicy(private_network_allowlist=["not-an-ip"])

    # --- Local files --------------------------------------------------- #
    def test_default_blocks_local_files(self):
        file_url = "file://localhost/" + self.secret.replace("\\", "/").lstrip("/")
        self.assertIsNone(resources.fetch_external_css(file_url))
        self.assertIsNone(resources.fetch_external_css(self.secret))
        self.assertIsNone(resources.fetch_image_data(TEST_IMG))

    def test_local_files_confined_to_root(self):
        policy = resources.ResourcePolicy(local_files_root=test_dir)
        self.assertIsNotNone(resources.fetch_image_data(TEST_IMG, policy))
        self.assertIsNotNone(resources.fetch_image_data("assets/images/test_img.png", policy))
        self.assertIsNone(resources.fetch_external_css(self.secret, policy))
        self.assertIsNone(resources.fetch_external_css("../../../../../../etc/passwd", policy))
        self.assertIsNone(resources.fetch_external_css("file:///etc/passwd", policy))
        self.assertIsNone(resources.fetch_external_css("file://otherhost/etc/passwd", policy))

    def test_local_root_via_options(self):
        parser = HtmlToDocx()
        parser.options["local-files-root"] = test_dir
        doc = parser.parse_html_string(f'<img src="{TEST_IMG}">')
        self.assertEqual(len(doc.inline_shapes), 1)

        parser = HtmlToDocx()
        doc = parser.parse_html_string(f'<img src="{TEST_IMG}">')
        self.assertEqual(len(doc.inline_shapes), 0)


if __name__ == "__main__":
    unittest.main()
