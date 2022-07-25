from odf.opendocument import OpenDocumentText, OpenDocumentPresentation
from odf.style import Style, TextProperties
from odf.presentation import Header
from odf.text import Title, H


class Presentation(object):
    def __init__(self):
        self.doc = OpenDocumentPresentation()

    def save(self, filename: str) -> None:
        self.doc.save(outputfile=filename)

    def add_header(self, header_text: str):
        # odf.element.IllegalChild: <text:title> is not allowed in <office:presentation>
        # h = Title(text=header_text)

        # odf.element.IllegalChild: <presentation:header> is not allowed in <office:presentation>
        # h = Header()

        # odf.element.IllegalChild: <text:h> is not allowed in <office:presentation>
        # h = H(text=header_text, outlinelevel=1)

        # self.doc.body.addElement(h)
        self.doc.presentation.addElement(h)


def main():
    p = Presentation()
    p.add_header("OMG LOL")
    p.save(filename="test.odp")


if __name__ == "__main__":
    main()
