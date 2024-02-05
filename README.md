# HTML FOR DOCX
Convert html to docx, this project is a fork from descontinued [pqzx/html2docx](https://github.com/pqzx/html2docx).

### To install

`pip install html-for-docx`

### Usage

Take a look on [pqzx/html2docx](https://github.com/pqzx/html2docx) project to see examples of usage.

### Why

My goal to fork and fix/update this package was to complete my current task at work that envolves manipulating a html to docs which the original couldn't complete because was lacking of few features and bugs, so instead creating a package from zero, I prefer update this one.

### Differences (fixes and new features)

**Fixes**
- Handle missing run for leading br tag | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/53)
- Fix base64 images | [djplaner](https://github.com/djplaner) from [Issue](https://github.com/pqzx/html2docx/issues/28#issuecomment-1052736896)
- Handle img tag without src attribute | [johnjor](https://github.com/johnjor) from [PR](https://github.com/pqzx/html2docx/pull/63)
- Fix text-align bug when `!important` | [Dfop02](https://github.com/dfop02)
- Fix background-color always set default color | [Dfop02](https://github.com/dfop02)
- Fix 'style lookup by style_id is deprecated.' | [Dfop02](https://github.com/dfop02)

**New Features**
- Add Witdh/Height style to images | [maifeeulasad](https://github.com/maifeeulasad) from [PR](https://github.com/pqzx/html2docx/pull/29)
- Support px, cm and % for style margin-left to paragraphs | [Dfop02](https://github.com/dfop02)
- Improve performance on large tables | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/58)
- Support for HTML Pagination | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support Table style | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support alternative encoding | [HebaElwazzan](https://github.com/HebaElwazzan) from [PR](https://github.com/pqzx/html2docx/pull/59)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
