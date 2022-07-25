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

        self.titlestyle = Style(name="MyMaster-title", family="presentation")
        self.titlestyle.addElement(ParagraphProperties(textalign="center"))
        self.titlestyle.addElement(TextProperties(fontsize="34pt"))
        self.titlestyle.addElement(GraphicProperties(fillcolor="#ffff99"))
        self.doc.styles.addElement(self.titlestyle)

        self.photostyle = Style(name="MyMaster-photo", family="presentation")
        self.doc.styles.addElement(self.photostyle)

        self.dpstyle = Style(name="dp1", family="drawing-page")
        self.dpstyle.addElement(
            DrawingPageProperties(
                transitiontype="automatic",
                transitionstyle="move-from-top",
                duration="PT05S",
            )
        )
        self.doc.automaticstyles.addElement(self.dpstyle)

        self.masterpage = MasterPage(name="MyMaster", pagelayoutname=pagelayout)
        self.doc.masterstyles.addElement(self.masterpage)

    def save(self, filename: str) -> None:
        self.doc.save(outputfile=filename)

    def add_page(self, title: str):
        page = Page(stylename=self.dpstyle, masterpagename=self.masterpage)
        self.doc.presentation.addElement(page)
        titleframe = Frame(
            stylename=self.titlestyle, width="25cm", height="2cm", x="1.5cm", y="0.5cm"
        )

        textbox = TextBox()
        titleframe.addElement(textbox)
        textbox.addElement(P(text=title))
        page.addElement(titleframe)

        # photoframe = Frame(stylename=photostyle, width="25cm", height="18.75cm", x="1.5cm", y="2.5cm")
        # page.addElement(photoframe)
        # href = self.doc.addPicture(picture[0])
        # photoframe.addElement(Image(href=href))


def main():
    p = Presentation()
    for word in "oh cool wow so amazing".split():
        p.add_page(word)
    p.save(filename="test.odp")


if __name__ == "__main__":
    main()
