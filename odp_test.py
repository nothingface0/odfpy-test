import lorem
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

        # Style for page titles
        self.titlestyle = Style(name="MyMaster-title", family="presentation")
        self.titlestyle.addElement(ParagraphProperties(textalign="center"))
        self.titlestyle.addElement(TextProperties(fontsize="34pt"))
        self.titlestyle.addElement(GraphicProperties(fillcolor="#ffff99"))
        self.doc.styles.addElement(self.titlestyle)

        # Style for adding content
        self.contentstyle = Style(name="MyMaster-content", family="presentation")
        self.contentstyle.addElement(ParagraphProperties(textalign="left"))
        self.contentstyle.addElement(TextProperties(fontsize="12pt"))
        self.contentstyle.addElement(GraphicProperties(fillcolor="#000000"))
        self.doc.styles.addElement(self.contentstyle)

        # Style for images
        self.photostyle = Style(name="MyMaster-photo", family="presentation")
        self.doc.styles.addElement(self.photostyle)

        # Style for pages and transitions
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

    def add_page(self, title: str, content: str):
        page = Page(stylename=self.dpstyle, masterpagename=self.masterpage)
        self.doc.presentation.addElement(page)

        titleframe = Frame(
            stylename=self.titlestyle, width="25cm", height="2cm", x="1.5cm", y="0.5cm"
        )
        textbox = TextBox()
        titleframe.addElement(textbox)
        textbox.addElement(P(text=title))
        page.addElement(titleframe)

        if content:
            c_text = TextBox()
            c_text.addElement(P(text=content))
            c_frame = Frame(
                stylename=self.contentstyle,
                width="25cm",
                height="4cm",
                x="1.5cm",
                y="7cm",
            )
            c_frame.addElement(c_text)
            page.addElement(c_frame)

        # photoframe = Frame(stylename=photostyle, width="25cm", height="18.75cm", x="1.5cm", y="2.5cm")
        # page.addElement(photoframe)
        # href = self.doc.addPicture(picture[0])
        # photoframe.addElement(Image(href=href))


def main():
    p = Presentation()
    for word in "What's next for Brussels, pooping eagles?".split():
        p.add_page(title=word, content=next(lorem.paragraph(count=1)))
    p.save(filename="test.odp")


if __name__ == "__main__":
    main()
