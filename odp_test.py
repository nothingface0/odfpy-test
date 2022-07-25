from odf.opendocument import OpenDocumentText, OpenDocumentPresentation
from odf.style import (
    Style,
    MasterPage,
    PageLayout,
    PageLayoutProperties,
    TextProperties,
    GraphicProperties,
    ParagraphProperties,
    DrawingPageProperties,
)
from odf.text import P
from odf.presentation import Header
from odf.text import Title, H
from odf.draw import Page, Frame, TextBox, Image


class Presentation(object):
    def __init__(self):
        self.doc = OpenDocumentPresentation()

        pagelayout = PageLayout(name="MyLayout")
        self.doc.automaticstyles.addElement(pagelayout)
        pagelayout.addElement(
            PageLayoutProperties(
                margin="0cm",
                pagewidth="28cm",
                pageheight="21cm",
                printorientation="landscape",
            )
        )

        titlestyle = Style(name="MyMaster-title", family="presentation")
        titlestyle.addElement(ParagraphProperties(textalign="center"))
        titlestyle.addElement(TextProperties(fontsize="34pt"))
        titlestyle.addElement(GraphicProperties(fillcolor="#ffff99"))
        self.doc.styles.addElement(titlestyle)

        photostyle = Style(name="MyMaster-photo", family="presentation")
        self.doc.styles.addElement(photostyle)

        masterpage = MasterPage(name="MyMaster", pagelayoutname=pagelayout)
        self.doc.masterstyles.addElement(masterpage)

    def save(self, filename: str) -> None:
        self.doc.save(outputfile=filename)

    def add_header(self, header_text: str):
        # odf.element.IllegalChild: <text:title> is not allowed in <office:presentation>
        h = Title(text=header_text)

        # odf.element.IllegalChild: <presentation:header> is not allowed in <office:presentation>
        # h = Header()

        # odf.element.IllegalChild: <text:h> is not allowed in <office:presentation>
        # h = H(text=header_text, outlinelevel=1)

        # self.doc.body.addElement(h)
        self.doc.presentation.addElement(h)


def main():
    p = Presentation()
    # p.add_header("OMG LOL")
    p.save(filename="test.odp")


if __name__ == "__main__":
    main()
