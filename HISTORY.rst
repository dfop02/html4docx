.. :changelog:

Release History
---------------

1.1.2 (2025-12-07)
++++++++++++++++++

**Updates**

- Update Pypi Workflow to Validate Tag, Check Lint, Run tests, Build distribution, Publish to PyPI and create release automatically. | `dfop02 <https://github.com/dfop02>`_

**Fixes**

- Fix text decoration color not being applied correctly. | `dfop02 <https://github.com/dfop02>`_
- Fix HISTORY.rst formatting. | `dfop02 <https://github.com/dfop02>`_

**New Features**

- [`PR #44 <https://github.com/dfop02/html4docx/pull/44>`_] Added support for custom css class to word style mappings. | `raithedavion <https://github.com/raithedavion>`_
- [`PR #44 <https://github.com/dfop02/html4docx/pull/44>`_] Added support for html tag to style overrides. | `raithedavion <https://github.com/raithedavion>`_
- [`PR #44 <https://github.com/dfop02/html4docx/pull/44>`_] Added support for setting default word style for new documents. | `raithedavion <https://github.com/raithedavion>`_
- [`PR #44 <https://github.com/dfop02/html4docx/pull/44>`_] Added support for "!important" style precedence. | `raithedavion <https://github.com/raithedavion>`_
- Change linter from flake8 to ruff. | `dfop02 <https://github.com/dfop02>`_


1.1.1 (2025-11-26)
++++++++++++++++++

**Updates**

- Fix Pypi Workflow.
- [`PR #48 <https://github.com/dfop02/html4docx/pull/48>`_] Add support for common CSS properties on text for: ``<h>``, ``<p>`` and ``<span>`` | `Lynuxen <https://github.com/Lynuxen>`_

**Fixes**

- Fixes `#46 <https://github.com/dfop02/html4docx/issues/46>`_: background-color style property highlights the whole paragraph instead of a single word
- Fixes `#47 <https://github.com/dfop02/html4docx/issues/47>`_: text-decoration style property for underline is not applied

**New Features**

- Add support for HTML Comments. | `Dfop02 <https://github.com/dfop02>`_
- Add support for text-align,line-height, margin-left, margin-right, text-indent for paragraphs
- Add support for the following text properties (applies to ``<span>``, ``<p>`` and ``<h>`` tags):
    - font-weight: ('bold', 'bolder', '700', '800', '900', 'normal', 'lighter', '400', '300', '100')
    - font-style: ('italic', 'oblique', 'normal'')
    - text-decoration: ('underline', 'line-through') ('solid', 'double', 'dotted', 'dashed', 'wavy'), and the longhand properties (text-decoration-\*)
    - text-transform: ('uppercase', 'lowercase', 'capitalize')
    - font-size
    - font-family
    - color
    - background-color: Paragraph and run highlight colors can now differ. Partial support on what can be used as a color.


1.1.0 (2025-11-01)
++++++++++++++++++

Major Changes
-------------

**Updates**

- Start Modularization of HtmlToDocx class.
- Add Typing hint for few relevant methods.

**Fixes**

- Fixed skip table was not working correctly.
- Fixed bug on styles parsing when style contains multiple colon.
- `PR #40 <https://github.com/dfop02/html4docx/issues/40>`_ Fixed styles was only being applied to span tag.

**New Features**

- Add built-in save for save docx
- Able to save in memory (BytesIO)
- Support to Python 3.14


1.0.10 (2025-08-20)
+++++++++++++++++++

**Updates**

- Update tests for all new scenarios. | `Dfop02 <https://github.com/dfop02>`_ and `hxzrx <https://github.com/hxzrx>`_

**Fixes**

- Fix rowspan and colspan for complex tables | `hxzrx <https://github.com/hxzrx>`_ from `PR #33 <https://github.com/dfop02/html4docx/pull/33>`_

**New Features**

- Add support for border with more complex values and keywords | `Dfop02 <https://github.com/dfop02>`_ from `Issue #31 <https://github.com/dfop02/html4docx/issues/31>`_


1.0.9 (2025-07-18)
++++++++++++++++++

**Updates**

- Starting modularize project with metadata. | `Dfop02 <https://github.com/dfop02>`_
- Update tests, removing useless tests and separating by modules. | `Dfop02 <https://github.com/dfop02>`_

**Fixes**

- Merge missing ``Release/1.0.8`` features. | `Dfop02 <https://github.com/dfop02>`_

**New Features**

- Add support for rowspan and colspan in tables. | `Dfop02 <https://github.com/dfop02>`_ from `Issue #25 <https://github.com/dfop02/html4docx/issues/25>`_
- Add support for 'vertical-align' in table cells. | `Dfop02 <https://github.com/dfop02>`_
- Add support for metadata | `Dfop02 <https://github.com/dfop02>`_


1.0.8 (2025-07-04)
++++++++++++++++++

**Updates**

- Add tests for image without paragraph. | `Dfop02 <https://github.com/dfop02>`_
- Add tests for bookmark without paragraph. | `Dfop02 <https://github.com/dfop02>`_
- Add tests for local image. | `Dfop02 <https://github.com/dfop02>`_
- Add tests for unbalanced table. | `Dfop02 <https://github.com/dfop02>`_

**Fixes**

- Fix crash when there is bookmark without paragraph. | `Dfop02 <https://github.com/dfop02>`_ from `Issue #21 <https://github.com/dfop02/html4docx/issues/21>`_
- Fix crash when there is image without paragraph. | `Dfop02 <https://github.com/dfop02>`_ from `Issue #19 <https://github.com/dfop02/html4docx/issues/19>`_

**New Features**

None


1.0.7 (2025-06-17)
++++++++++++++++++

**Updates**

- Add tests for inline images. | `Dfop02 <https://github.com/dfop02>`_
- Add tests for Bold, Italic, Underline and Strike. | `Dfop02 <https://github.com/dfop02>`_
- Add tests for Ordered and Unordered Lists. | `TaylorN15 <https://github.com/TaylorN15>`_ from `PR #16 <https://github.com/dfop02/html4docx/pull/16>`_
- Update Docs (Include Known Issues and Project Guidelines). | `Dfop02 <https://github.com/dfop02>`_
- Refactor ``utils.py`` file. | `Dfop02 <https://github.com/dfop02>`_

**Fixes**

- Fix Ordered and Unordered Lists. | `TaylorN15 <https://github.com/TaylorN15>`_ from `PR #16 <https://github.com/dfop02/html4docx/pull/16>`_

**New Features**

- Being able to use inline images on same paragraph. | `Dfop02 <https://github.com/dfop02>`_
- Limit 5s timeout to fetch any image from web. | `Dfop02 <https://github.com/dfop02>`_


1.0.6 (2025-05-02)
++++++++++++++++++

**Updates**

- Fix Changelog bad formating.
- Update ``README.md`` with latest changes.
- Add funding to project.
- Add ``CONTRIBUTING.md`` to project.
- Add pull request template to project.
- Update tests for table style to assert that is working fine.
- Save ``tests.docx`` on Github Actions to make it easier to help debugging issues across multiple builds. | `gionn <https://github.com/gionn>`_

**Fixes**

- Fix ``table_style`` not working. | `Dfop02 <https://github.com/dfop02>`_


1.0.5 (2025-02-23)
++++++++++++++++++

**Updates**

- Refactory functions to be more readable and performatic, moving common functions to utils.
- Add tests for table cells new features.

**Fixes**

- Fix a bug when using colors by name, when some colors exists but was not available. E.g.: magenta | `Dfop02 <https://github.com/dfop02>`_
- Fix background-color not working, always returning black. | `Dfop02 <https://github.com/dfop02>`_

**New Features**

- Add support to table cells style (border, background-color, width, height, margin) | `Dfop02 <https://github.com/dfop02>`_
- Add support to "in", "rem", "em", "mm" and "pc" units | `Dfop02 <https://github.com/dfop02>`_


1.0.4 (2024-08-11)
++++++++++++++++++

**Updates**

- Create Changelog HISTORY.
- Update README.
- Add Github Action Workflow to publish in pypi.
- Change default VERSION tag, removing the "v" from new releases.

**New Features**

- Support to internal links (Anchor) | `Dfop02 <https://github.com/dfop02>`_


1.0.3 (2024-02-27)
++++++++++++++++++

- Adapt font_size when text, ex.: small, medium, etc. | `Dfop02 <https://github.com/dfop02>`_
- Fix error for image weight and height when no digits | `Dfop02 <https://github.com/dfop02>`_


1.0.2 (2024-02-20)
++++++++++++++++++

- Support px, cm, pt and % for style margin-left to paragraphs | `Dfop02 <https://github.com/dfop02>`_
- Fix 'style lookup by style_id is deprecated.' | `Dfop02 <https://github.com/dfop02>`_
- Fix bug when any style has ``!important`` | `Dfop02 <https://github.com/dfop02>`_
- Refactory Tests to be more consistent and less 'human validation' | `Dfop02 <https://github.com/dfop02>`_
- Support to color by name | `Dfop02 <https://github.com/dfop02>`_


1.0.1 (2024-02-05)
++++++++++++++++++

- Fix README.


1.0.0 (2024-02-05)
+++++++++++++++++++

- Initial Release!

**Fixes**

- Handle missing run for leading br tag | `dashingdove <https://github.com/dashingdove>`_ from `PR #53 <https://github.com/pqzx/html2docx/pull/53>`_
- Fix base64 images | `djplaner <https://github.com/djplaner>`_ from `Issue #28 <https://github.com/pqzx/html2docx/issues/28#issuecomment-1052736896>`_
- Handle img tag without src attribute | `johnjor <https://github.com/johnjor>`_ from `PR #63 <https://github.com/pqzx/html2docx/pull/63>`_

**New Features**

- Add Witdh/Height style to images | `maifeeulasad <https://github.com/maifeeulasad>`_ from `PR #29 <https://github.com/pqzx/html2docx/pull/29>`_
- Improve performance on large tables | `dashingdove <https://github.com/dashingdove>`_ from `PR #58 <https://github.com/pqzx/html2docx/pull/58>`_
- Support for HTML Pagination | `Evilran <https://github.com/Evilran>`_ from `PR #39 <https://github.com/pqzx/html2docx/pull/39>`_
- Support Table style | `Evilran <https://github.com/Evilran>`_ from `PR #39 <https://github.com/pqzx/html2docx/pull/39>`_
- Support alternative encoding | `HebaElwazzan <https://github.com/HebaElwazzan>`_ from `PR #59 <https://github.com/pqzx/html2docx/pull/59>`_
