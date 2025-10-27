.. :changelog:

Release History
---------------

1.1.0 (2025-10-xx)
++++++++++++++++++

Major Changes
-------------

**Updates**

- Start Modularization of HtmlToDocx class.
- Add Typing hint for few relevant methods.

**Fixes**

- Fixed skip table was not working correctly.
- Fixed bug on styles parsing when style contains multiple colon.
- [PR](https://github.com/dfop02/html4docx/issues/40) Fixed styles was only being applied to span tag.

**New Features**

- Add built-in save for save docx
- Able to save in memory (BytesIO)


1.0.10 (2025-08-20)
++++++++++++++++++

**Updates**

- Update tests for all new scenarios. | [Dfop02](https://github.com/dfop02) and [hxzrx](https://github.com/hxzrx)

**Fixes**

- Fix rowspan and colspan for complex tables | [hxzrx](https://github.com/hxzrx) from [PR](https://github.com/dfop02/html4docx/pull/33)

**New Features**

- Add support for border with more complex values and keywords | [Dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/31)


1.0.9 (2025-07-18)
++++++++++++++++++

**Updates**

- Starting modularize project with metadata. | [Dfop02](https://github.com/dfop02)
- Update tests, removing useless tests and separating by modules. | [Dfop02](https://github.com/dfop02)

**Fixes**

- Merge missing `Release/1.0.8` features. | [Dfop02](https://github.com/dfop02)

**New Features**

- Add support for rowspan and colspan in tables. | [Dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/25)
- Add support for 'vertical-align' in table cells. | [Dfop02](https://github.com/dfop02)
- Add support for metadata | [Dfop02](https://github.com/dfop02)


1.0.8 (2025-07-04)
++++++++++++++++++

**Updates**

- Add tests for image without paragraph. | [Dfop02](https://github.com/dfop02)
- Add tests for bookmark without paragraph. | [Dfop02](https://github.com/dfop02)
- Add tests for local image. | [Dfop02](https://github.com/dfop02)
- Add tests for unbalanced table. | [Dfop02](https://github.com/dfop02)

**Fixes**

- Fix crash when there is bookmark without paragraph. | [Dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/21)
- Fix crash when there is image without paragraph. | [Dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/19)

**New Features**

None


1.0.7 (2025-06-17)
++++++++++++++++++

**Updates**

- Add tests for inline images. | [Dfop02](https://github.com/dfop02)
- Add tests for Bold, Italic, Underline and Strike. | [Dfop02](https://github.com/dfop02)
- Add tests for Ordered and Unordered Lists. | [TaylorN15](https://github.com/TaylorN15) from [PR](https://github.com/dfop02/html4docx/pull/16)
- Update Docs (Include Known Issues and Project Guidelines). | [Dfop02](https://github.com/dfop02)
- Refactor `utils.py` file. | [Dfop02](https://github.com/dfop02)

**Fixes**

- Fix Ordered and Unordered Lists. | [TaylorN15](https://github.com/TaylorN15) from [PR](https://github.com/dfop02/html4docx/pull/16)

**New Features**

- Being able to use inline images on same paragraph. | [Dfop02](https://github.com/dfop02)
- Limit 5s timeout to fetch any image from web. | [Dfop02](https://github.com/dfop02)


1.0.6 (2025-05-02)
++++++++++++++++++

**Updates**

- Fix Changelog bad formating.
- Update `README.md` with latest changes.
- Add funding to project.
- Add `CONTRIBUTING.md` to project.
- Add pull request template to project.
- Update tests for table style to assert that is working fine.
- Save `tests.docx` on Github Actions to make it easier to help debugging issues across multiple builds. | [gionn](https://github.com/gionn)

**Fixes**

- Fix `table_style` not working. | [Dfop02](https://github.com/dfop02)


1.0.5 (2025-02-23)
++++++++++++++++++

**Updates**

- Refactory functions to be more readable and performatic, moving common functions to utils.
- Add tests for table cells new features.

**Fixes**

- Fix a bug when using colors by name, when some colors exists but was not available. E.g.: magenta | [Dfop02](https://github.com/dfop02)
- Fix background-color not working, always returning black. | [Dfop02](https://github.com/dfop02)

**New Features**

- Add support to table cells style (border, background-color, width, height, margin) | [Dfop02](https://github.com/dfop02)
- Add support to "in", "rem", "em", "mm" and "pc" units | [Dfop02](https://github.com/dfop02)


1.0.4 (2024-08-11)
++++++++++++++++++

**Updates**

- Create Changelog HISTORY.
- Update README.
- Add Github Action Workflow to publish in pypi.
- Change default VERSION tag, removing the "v" from new releases.

**New Features**

- Support to internal links (Anchor) | [Dfop02](https://github.com/dfop02)


1.0.3 (2024-02-27)
++++++++++++++++++

- Adapt font_size when text, ex.: small, medium, etc. | [Dfop02](https://github.com/dfop02)
- Fix error for image weight and height when no digits | [Dfop02](https://github.com/dfop02)


1.0.2 (2024-02-20)
++++++++++++++++++

- Support px, cm, pt and % for style margin-left to paragraphs | [Dfop02](https://github.com/dfop02)
- Fix 'style lookup by style_id is deprecated.' | [Dfop02](https://github.com/dfop02)
- Fix bug when any style has `!important` | [Dfop02](https://github.com/dfop02)
- Refactory Tests to be more consistent and less 'human validation' | [Dfop02](https://github.com/dfop02)
- Support to color by name | [Dfop02](https://github.com/dfop02)


1.0.1 (2024-02-05)
++++++++++++++++++

- Fix README.


1.0.0 (2024-02-05)
+++++++++++++++++++

- Initial Release!

**Fixes**

- Handle missing run for leading br tag | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/53)
- Fix base64 images | [djplaner](https://github.com/djplaner) from [Issue](https://github.com/pqzx/html2docx/issues/28#issuecomment-1052736896)
- Handle img tag without src attribute | [johnjor](https://github.com/johnjor) from [PR](https://github.com/pqzx/html2docx/pull/63)

**New Features**

- Add Witdh/Height style to images | [maifeeulasad](https://github.com/maifeeulasad) from [PR](https://github.com/pqzx/html2docx/pull/29)
- Improve performance on large tables | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/58)
- Support for HTML Pagination | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support Table style | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support alternative encoding | [HebaElwazzan](https://github.com/HebaElwazzan) from [PR](https://github.com/pqzx/html2docx/pull/59)
