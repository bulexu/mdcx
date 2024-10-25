import sys
import docx
from pathlib import Path


def _style_title_border(style_title):
    """Removes border style on title which is set by python-docx by default.
    This is a hack because there's no programmatic way to do this as of writing"""
    el = style_title._element
    el.remove(el.xpath("w:pPr")[0])


def _rm_toc(md: str) -> list:
    """Removes first table of contents section from a markdown string, returning list of lines"""
    # Don't check if there isn't one
    check = md.lower()
    if "table of contents" not in check and "contents" not in check:
        return md.splitlines()
    # Parse through
    in_toc = False
    removed_toc = False
    keep = []
    for line in md.splitlines():
        clean = line.lstrip()
        # Title, so either start/end toc removal
        if clean.startswith("#") and not removed_toc:
            # Stop removing toc
            if in_toc:
                in_toc = False
                keep.append(line)
                continue
            # Start removing toc
            title = clean.lstrip("#").strip().lower()
            if title in ["table of contents", "contents"]:
                in_toc = True
            else:
                keep.append(line)
        # Add like normal
        elif not in_toc:
            keep.append(line)
    return keep



def _is_bib(text: str) -> bool:
    """Checks if provided heading text is referencing a bibliography"""
    return text.lower() in ["bibliography", "references"]


def _level_info(line: str) -> tuple:
    """Figures out level information and returns it and the line without spacing"""
    stripped = line.lstrip()
    num = len(line) - len(stripped)
    level = int(num / 2)
    return (level, stripped)


def _err_exit(msg: str):
    """Prints error message to console and exits program, used for command-line"""
    print(f"Error: {msg}", file=sys.stderr)
    sys.exit(1)


def _add_link(
    paragraph: docx.text.paragraph.Paragraph, link: str, text: str, external: bool
):
    """Places an internal or external link within a paragraph object"""

    # Create the w:hyperlink tag
    hyperlink = docx.oxml.shared.OxmlElement("w:hyperlink")
    # Set where it links to
    if external:
        # This gets access to the document.xml.rels file and gets a new relation id value
        part = paragraph.part
        r_id = part.relate_to(
            link, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True
        )
        # External relationship value
        hyperlink.set(docx.oxml.shared.qn("r:id"), r_id)
    else:
        # Internal anchor value
        hyperlink.set(docx.oxml.shared.qn("w:anchor"), link)
    # Create a w:r element
    new_run = docx.oxml.shared.OxmlElement("w:r")
    # Create a new w:rPr element
    run_prop = docx.oxml.shared.OxmlElement("w:rPr")
    # Add link styling
    run_style = docx.oxml.shared.OxmlElement("w:pStyle")
    run_style.set(docx.oxml.shared.qn("w:val"), "Link")
    run_prop.append(run_style)
    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(run_prop)
    new_run.text = text
    hyperlink.append(new_run)
    # Add to paragraph
    paragraph._p.append(hyperlink)
    return hyperlink


def _is_bib(text: str) -> bool:
    """Checks if provided heading text is referencing a bibliography"""
    return text.lower() in ["bibliography", "references"]


def get_docx_path(args: list[str], md_path: Path) -> Path:
    # Provide just normal if it's there
    if len(args) > 1:
        return Path(args[1])

    # Base if on first arg if not
    return Path.cwd() / Path(md_path.stem + ".docx")
