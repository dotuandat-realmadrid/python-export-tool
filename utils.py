import textwrap

def wrap_text(text, width=50):
    """Bọc văn bản để hiển thị trên nhiều dòng."""
    if text is None:
        return ""
    return '\n'.join(textwrap.wrap(text, width=width))