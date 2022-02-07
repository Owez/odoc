from pathlib import Path
from docx import Document
import re


class ODoc:
    def __init__(self, md: str, verbose: bool = True) -> None:
        # add basics
        self.lines = md.splitlines()
        self.doc = Document()
        self.verbose = verbose

        # parse
        self._parse()

    def save(self, path: Path):
        self.doc.save(path)

    def _info(self, msg: str, tabbing: str = "  "):
        if self.verbose:
            print(f"{tabbing}{msg}..")

    def _parse(self):
        self._info("Generating document", "")

        # line parsing
        for ind, line in enumerate(self.lines):

            def is_indent() -> bool:
                return (
                    line.startswith(" ")
                    or line.startswith("\t")
                    or line.startswith("-")
                    or str_is_list_item(line)
                )

            if len(line) == 0:
                # empty line, skip
                continue
            elif line.startswith("#"):
                # heading
                self._add_heading(line)
            elif is_indent():
                # start intent parsing
                self._parse_indents(ind)
            else:
                # paragraph
                self._add_paragraph(line)

        # add header numbering
        _add_heading_numbering(self.doc)

    def _parse_indents(self, from_ind: int):
        def get_line(lines: list, ind: int) -> str:
            return

        def get_level(line, spaces: bool = None) -> tuple:
            if spaces is None:
                # unknown levelling
                count_s = _count_chars(line, " ")
                count_t = _count_chars(line, "\t")
                return (max(count_s, count_t), count_s >= count_t)
            elif spaces:
                # space levelling
                return (_count_chars(line, " "), spaces)
            else:
                # tab levelling
                pass

        lines = self.lines[from_ind:]
        ind = 0

        spaces = None

        while True:
            line = lines[ind] if len(lines) > ind else None
            if line is None:
                return

            (level,spaces) = get_level(line,spaces)
            
            # TODO: use level to get same paragraphs on next line and then add to bulletpoint or list item to completion

    def _add_heading(self, line: str):
        level = _count_chars(line, "#")
        text = line[level:].strip()

        self._info(f"Adding heading (" + "#" * (level + 1) + ")")
        self.doc.add_heading(text, level)

    def _add_bulletpoint(self, line: str, level: int):
        self._info(f"Adding level {level} bulletpoint")

        indenting = "\t" * level
        self.doc.add_paragraph(indenting, "List Bullet")

        text = line[1:].strip()
        self._add_paragraph_runs(text)

    def _add_listpoint(self, line: str, level: int):
        self._info(f"Adding level {level} listpoint ({line[0]}.)")

        indenting = "\t" * level
        self.doc.add_paragraph(indenting, "List Number")

        text = line[2:].strip()
        self._add_paragraph_runs(text)

    def _add_paragraph(self, line):
        self._info("Adding paragraph")

        paragraph = self.doc.add_paragraph()
        self._add_paragraph_runs(line, paragraph)

    def _add_paragraph_runs(self, line, paragraph):
        self._info("Adding next paragraph run", "    ")

        paragraph.add_run(line)  # TODO: parse line
        pass  # TODO: parse paragraph bits


def _iter_heading(paragraphs):
    for paragraph in paragraphs:
        isItHeading = re.match("Heading ([1-9])", paragraph.style.name)
        if isItHeading:
            yield int(isItHeading.groups()[0]), paragraph


def _add_heading_numbering(document):
    hNums = [0, 0, 0, 0, 0]
    for index, hx in _iter_heading(document.paragraphs):
        # ---put zeroes below---
        for i in range(index + 1, 5):
            hNums[i] = 0
        # ---increment this---
        hNums[index] += 1
        # ---prepare the string---
        hStr = ""
        for i in range(1, index + 1):
            hStr += "%d." % hNums[i]
        # ---add the numbering---
        hx.text = hStr + " " + hx.text


def _count_chars(line: str, target: str) -> int:
    count = 0
    for c in line:
        if c == target:
            count += 1
        else:
            break
    return count


def str_is_list_item(string: str) -> bool:
    return len(string) > 1 and string[0].isnumeric() and string[1] == "."
