# mdcx

Seamless markdown to docx converter.

## Example

Command-line:

```shell
$ poetry run python main.py examples/test.md test.docx
```

In Python:

```python
from pathlib import Path
from src.document import Document
from src.styles import Style

# 从文件读取 Markdown 内容
md_path = Path("path/to/your/markdown/file.md")
with open(md_path, "r", encoding="utf-8") as file:
    md_content = file.read()

# 创建 Document 对象
style = Style.andy()  # 或者使用 Style.foxtrot()
doc = Document(md_path, style)

# 保存为 docx 文件
output_path = Path("output.docx")
doc.save(output_path)
```

## Installation

### Init
```shell
$ pip install poetry
$ poetry install
```

## Showcase

Here's a generated document from the `examples/` directory using the default theme:

![AirBnB Document](examples/images/airbnb.png)

## Roadmap

Here are the upcoming features for the development of mdcx:

- Markdown:
  - [ ] Heading links
  - [ ] Tables
- Quality-of-life
  - [ ] Support `#` titles as well as the current yml titles
  - [ ] Support a basic version of TOML `+++` metadata
- Extras:
  - [ ] Local URIs become automatic managed appendixes

This project isn't finished as not all basic markdown has been implemented. The hope for this project is to be able to seamlessly convert all well-formatted markdown to a docx.
