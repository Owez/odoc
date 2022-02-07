from odoc import ODoc

md = """# Hello world

This is a little paragraph with nothing really on it

- One
- Two

## Another Section

1. Three
2. Four

### Final bit

Paragraph"""

odoc = ODoc(md)
odoc.save("output.docx")
