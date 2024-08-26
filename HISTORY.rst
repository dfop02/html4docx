.. :changelog:

Release History
---------------

1.0.5 (2024-08-21)
++++++++++++++++++

**Fixes**
- Fix numeric lists | [Dfop02](https://github.com/dfop02) based on [Issue](https://github.com/python-openxml/python-docx/issues/25#issuecomment-400787031) from [madphysicist](https://github.com/madphysicist)
- Fix nested lists (numeric/bullet) | [Dfop02](https://github.com/dfop02)


1.0.4 (2024-08-06)
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
