
class Style:
    """Unified and modifiable style for a document"""

    def __init__(
        self,
        font_heading: str,
        font_body: str,
        font_code: str,
        body_pt: int,
        body_justified: bool,
        body_lines: float,
        heading_bold: bool,
        heading_blue: bool,
        # TODO: numbered headings
    ) -> None:
        self.font_heading = font_heading  # 标题字体
        self.font_body = font_body  # 正文字体
        self.font_code = font_code  # 代码字体
        self.body_pt = body_pt  # 正文字号
        self.body_justified = body_justified  # 正文是否两端对齐
        self.body_lines = body_lines  # 正文行距
        self.heading_bold = heading_bold  # 标题是否加粗
        self.heading_blue = heading_blue  # 标题是否为蓝色

    @staticmethod
    def andy():
        return Style("微软雅黑", "宋体", "仿宋", 12, False, 1.5, True, True)

    @staticmethod
    def foxtrot():
        return Style(
            "黑体",
            "宋体",
            "楷体",
            11,
            True,
            1.2,
            False,
            False,
        )

    def _body_alignment(self) -> int:
        return 3 if self.body_justified else 0
