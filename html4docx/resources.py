import base64
import binascii
import functools
import http.client
import ipaddress
import logging
import os
import socket
import urllib.error
import urllib.request
from io import BytesIO
from typing import Optional
from urllib.parse import urlparse

from html4docx import constants

logger = logging.getLogger(__name__)


class BlockedResourceError(Exception):
    """Raised when a resource is refused by the ResourcePolicy."""


class ResourcePolicy:
    def __init__(
        self,
        allow_remote: bool = True,
        allow_private_network: bool = False,
        private_network_allowlist=(),
        local_files_root: Optional[str] = None,
        max_bytes: int = constants.DEFAULT_MAX_RESOURCE_BYTES,
        timeout: float = 5,
    ):
        self.allow_remote = allow_remote
        self.allow_private_network = allow_private_network
        self.private_network_allowlist = _parse_allowlist(private_network_allowlist)
        self.local_files_root = local_files_root
        self.max_bytes = max_bytes
        self.timeout = timeout
        if self.allow_private_network and self.private_network_allowlist:
            logger.warning("'allow-private-network' is enabled, so 'private-network-allowlist' is ignored")

    @classmethod
    def from_options(cls, options: dict) -> "ResourcePolicy":
        return cls(
            allow_remote=options.get("allow-remote-resources", True),
            allow_private_network=options.get("allow-private-network", False),
            private_network_allowlist=options.get("private-network-allowlist") or (),
            local_files_root=options.get("local-files-root"),
            max_bytes=options.get("max-resource-bytes", constants.DEFAULT_MAX_RESOURCE_BYTES),
        )


def _parse_allowlist(entries) -> tuple:
    """Convert IPs/CIDRs to ip_network objects. Invalid entries raise ValueError (fail fast)."""
    if isinstance(entries, str):
        entries = [entries]
    return tuple(ipaddress.ip_network(str(entry).strip(), strict=False) for entry in entries)


def _is_forbidden_ip(ip_str: str, allowlist: tuple = ()) -> bool:
    ip = ipaddress.ip_address(ip_str.split("%", 1)[0])  # strip IPv6 zone id
    if ip.version == 6 and ip.ipv4_mapped:
        ip = ip.ipv4_mapped

    if any(ip in net for net in allowlist):
        return False
    return (not ip.is_global) or ip.is_multicast


def _guarded_create_connection(
    address, timeout=socket._GLOBAL_DEFAULT_TIMEOUT, source_address=None, *args, allowlist=(), **kwargs
):
    """
    Replacement for socket.create_connection that resolves the host once and
    connects only to vetted public IPs (no TOCTOU / DNS rebinding window).
    """
    host, port = address
    last_err = None
    for family, socktype, proto, _, sockaddr in socket.getaddrinfo(host, port, 0, socket.SOCK_STREAM):
        if _is_forbidden_ip(sockaddr[0], allowlist):
            last_err = BlockedResourceError(f"Refusing to connect to non-public address {sockaddr[0]} ({host})")
            continue
        sock = socket.socket(family, socktype, proto)
        try:
            if timeout is not socket._GLOBAL_DEFAULT_TIMEOUT:
                sock.settimeout(timeout)
            if source_address:
                sock.bind(source_address)
            sock.connect(sockaddr)
            return sock
        except OSError as err:
            sock.close()
            last_err = err
    raise last_err or OSError(f"Could not resolve {host}")


class _GuardedHTTPConnection(http.client.HTTPConnection):
    def __init__(self, *args, allowlist=(), **kwargs):
        super().__init__(*args, **kwargs)
        self._create_connection = functools.partial(_guarded_create_connection, allowlist=allowlist)


class _GuardedHTTPSConnection(http.client.HTTPSConnection):
    def __init__(self, *args, allowlist=(), **kwargs):
        super().__init__(*args, **kwargs)
        self._create_connection = functools.partial(_guarded_create_connection, allowlist=allowlist)


class _GuardedHTTPHandler(urllib.request.HTTPHandler):
    def __init__(self, allowlist=(), **kwargs):
        super().__init__(**kwargs)
        self._allowlist = allowlist

    def http_open(self, req):
        return self.do_open(_GuardedHTTPConnection, req, allowlist=self._allowlist)


class _GuardedHTTPSHandler(urllib.request.HTTPSHandler):
    def __init__(self, allowlist=(), **kwargs):
        super().__init__(**kwargs)
        self._allowlist = allowlist

    def https_open(self, req):
        return self.do_open(_GuardedHTTPSConnection, req, context=self._context, allowlist=self._allowlist)

def _build_opener(allow_private_network: bool, allowlist: tuple = ()) -> urllib.request.OpenerDirector:
    """
    Opener with ONLY http/https handlers: no file://, ftp://, data: handlers
    and no environment proxies, so redirects cannot escape to other schemes.
    """
    opener = urllib.request.OpenerDirector()
    if allow_private_network:
        handlers = [urllib.request.HTTPHandler(), urllib.request.HTTPSHandler()]
    else:
        handlers = [_GuardedHTTPHandler(allowlist=allowlist), _GuardedHTTPSHandler(allowlist=allowlist)]
    handlers += [
        urllib.request.HTTPDefaultErrorHandler(),
        urllib.request.HTTPRedirectHandler(),
        urllib.request.HTTPErrorProcessor(),
        urllib.request.UnknownHandler(),
    ]
    for handler in handlers:
        opener.add_handler(handler)
    return opener


def _fetch_remote_bytes(url: str, policy: ResourcePolicy) -> bytes:
    if not policy.allow_remote:
        raise BlockedResourceError("Remote resources are disabled ('allow-remote-resources')")
    if urlparse(url).scheme.lower() not in constants.REMOTE_RESOURCE_SCHEMES:
        raise BlockedResourceError(f"Scheme not allowed: {url}")
    opener = _build_opener(policy.allow_private_network, policy.private_network_allowlist)
    with opener.open(url, timeout=policy.timeout) as response:
        data = response.read(policy.max_bytes + 1)
    if len(data) > policy.max_bytes:
        raise BlockedResourceError(f"Resource larger than {policy.max_bytes} bytes: {url}")
    return data


def _resolve_local_path(src: str, policy: ResourcePolicy) -> str:
    """Resolve a local path / file:// URL, confined to policy.local_files_root."""
    if not policy.local_files_root:
        raise BlockedResourceError("Local files are disabled (set 'local-files-root' to enable)")

    parsed = urlparse(src)
    if parsed.scheme.lower() == "file":
        if parsed.netloc not in ("", "localhost"):
            raise BlockedResourceError(f"Remote file:// host not allowed: {src}")
        src = urllib.request.url2pathname(parsed.path)

    root = os.path.realpath(policy.local_files_root)
    candidate = os.path.realpath(os.path.join(root, src))  # absolute src replaces root, then gets checked
    try:
        inside = os.path.commonpath([root, candidate]) == root
    except ValueError:  # different drives on Windows
        inside = False
    if not inside:
        raise BlockedResourceError(f"Path outside of local-files-root: {src}")
    return candidate


def _read_local_bytes(src: str, policy: ResourcePolicy) -> bytes:
    path = _resolve_local_path(src, policy)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    if os.path.getsize(path) > policy.max_bytes:
        raise BlockedResourceError(f"File larger than {policy.max_bytes} bytes: {src}")
    with open(path, "rb") as f:
        return f.read()


def _classify_source(src: str) -> str:
    scheme = urlparse(src).scheme.lower()
    if scheme in constants.REMOTE_RESOURCE_SCHEMES:
        return "remote"
    if scheme == "file" or len(scheme) <= 1:  # "" or Windows drive letter (C:\...)
        return "local"
    return "blocked"


def fetch_resource_bytes(src: str, policy: Optional[ResourcePolicy] = None) -> Optional[bytes]:
    """
    Fetch a remote or local resource honoring the policy.
    Returns None (and logs a warning) if it's blocked or unavailable.
    """
    policy = policy or ResourcePolicy()
    try:
        kind = _classify_source(src)
        if kind == "remote":
            return _fetch_remote_bytes(src, policy)
        if kind == "local":
            return _read_local_bytes(src, policy)
        raise BlockedResourceError(f"Scheme not allowed: {src}")
    except BlockedResourceError as err:
        logger.warning("Blocked external resource: %s", err)
    except (urllib.error.URLError, http.client.HTTPException, OSError, ValueError) as err:
        logger.warning("Could not fetch external resource %s: %s", src, err)
    return None


def fetch_image(url: str, policy: Optional[ResourcePolicy] = None):
    """
    Fetches an image from a http(s) url honoring the ResourcePolicy.
    If successful returns a BytesIO object, else returns None.
    """
    if _classify_source(url) != "remote":
        logger.warning("Blocked external resource: not a http(s) url: %s", url)
        return None
    data = fetch_resource_bytes(url, policy)
    return BytesIO(data) if data is not None else None


def fetch_image_data(src: str, policy: Optional[ResourcePolicy] = None):
    """Fetches image data from a data URI, URL or (allowed) local file."""
    if src.startswith("data:image/"):  # Handle Base64
        try:
            _, encoded = src.split(",", 1)
            return BytesIO(base64.b64decode(encoded))
        except (ValueError, binascii.Error):
            return None

    data = fetch_resource_bytes(src, policy)
    return BytesIO(data) if data is not None else None


def fetch_external_css(url: str, policy: Optional[ResourcePolicy] = None) -> Optional[str]:
    """
    Fetch external CSS from URL or local file, honoring the ResourcePolicy.

    Returns:
        Optional[str]: CSS content as string, or None if blocked / fetch fails
    """
    if not url:
        return None

    content = fetch_resource_bytes(url, policy)
    if content is None:
        return None
    try:
        return content.decode("utf-8")
    except UnicodeDecodeError:
        return content.decode("latin-1")
