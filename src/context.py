from pathlib import Path
from .utils import _is_bib


class Context:
    """Contextual information for compartmentalised converting"""

    def __init__(self, wd: Path | None = None) -> None:
        self.line = 0
        self.heading = None
        self.italic = False
        self.bold = False
        self.underline = False
        self.strikethrough = False
        self.figures = 0
        self.wd = wd

    def no_spacing(self) -> bool:
        """Checks if elements should have spacing within the current section"""
        if self.heading is None:
            return False
        return _is_bib(self.heading.text)

    def next_line(self):
        """Skips to the next line"""
        self.line += 1
        self.char = 0
        self.italic = False
        self.bold = False
        self.underline = False
        self.strikethrough = False

    def flip_italic(self):
        """Flips italic value"""
        self.italic = not self.italic

    def flip_bold(self):
        """Flips bold value"""
        self.bold = not self.bold

    def link_to(self, link: str | Path) -> Path:
        """Gets link to something from the markdown file's directory"""
        return self.wd / link

