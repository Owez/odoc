from odoc import ODoc

md = """# Hello world

This is a little paragraph with nothing really on it

- One
    - Test
- Two

## Another Section

1. Three
2. Four

### Final bit

Paragraph"""

odoc = ODoc(md, True)
odoc.save("output.docx")
