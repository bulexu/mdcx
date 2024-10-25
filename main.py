import sys
from pathlib import Path

# 添加 src 目录到 Python 路径
sys.path.append(str(Path(__file__).parent / "src"))

from src.document import Document
from src.styles import Style
from src.utils import get_docx_path, _err_exit

CLI_HELP = """
使用方法: python -m src.main [in] [out] [options]
选项:
  --help     显示此帮助信息
  --foxtrot  使用 Foxtrot 样式
"""

def main():
    args = sys.argv[1:]
    if len(args) == 0:
        _err_exit("请提供 [in]")
    elif "--help" in args[2:]:
        print(CLI_HELP)
        sys.exit(0)

    foxtrot = "--foxtrot" in args[2:]
    md_path = Path(args[0])
    docx_path = get_docx_path(args, md_path)

    if not md_path.exists():
        raise Exception(f"Markdown 文件 '{args[0]}' 不存在")

    try:
        with open(md_path, "r", encoding="utf-8") as file:
            md = file.read()
    except Exception as e:
        _err_exit(f"Markdown 文件 '{args[0]}' 无效 ({e})")

    style = Style.andy() if not foxtrot else Style.foxtrot()
    Document(md, md_path, style).save(docx_path)

if __name__ == "__main__":
    main()
