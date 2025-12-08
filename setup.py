import os
from setuptools import setup, find_packages

here   = os.path.abspath(os.path.dirname(__file__))
README = open(os.path.join(here, 'README.md')).read()
VERSION = '1.1.2'

setup(
    name                 = 'html-for-docx',
    version              = VERSION,
    description          = 'Convert HTML to Docx easily and fastly',
    long_description     = README,
    license              = 'MIT',
    packages             = find_packages(),
    python_requires      = '>=3.7',
    author               = 'Diogo Fernandes',
    author_email         = 'diogofernandesop@gmail.com',
    platforms            = ['any'],
    include_package_data = True,
    keywords             = ['html', 'docx', 'docs', 'office', 'word', 'convert', 'transform'],
    zip_safe             = False,
    url                  = 'https://github.com/dfop02/html4docx',
    project_urls = {
        "Changelog": "https://github.com/dfop02/html4docx/blob/master/HISTORY.rst",
        "Bug Tracker": "https://github.com/dfop02/html4docx/issues",
        "Repository": "https://github.com/dfop02/html4docx"
    },
    download_url         = f'https://github.com/dfop02/html4docx/archive/{VERSION}.tar.gz',
    classifiers          = [
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Topic :: Software Development :: Build Tools',
        'Topic :: Software Development :: Libraries',
        'Topic :: Text Processing',
        'Topic :: Text Processing :: Markup :: HTML',
        'Topic :: Utilities',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Programming Language :: Python :: 3.13',
        'Programming Language :: Python :: 3.14'
    ],
    install_requires = [
        'python-docx>=1.1.0',
        'beautifulsoup4>=4.12.2'
    ]
)
