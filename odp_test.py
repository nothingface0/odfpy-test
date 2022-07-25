import lorem
from datetime import datetime
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
from odf import dc
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
        # The name attribute *must* start with the name of the masterpage
        # The name after the dash *also* matters (should be "title", "subtitle"
        # or "photo").
        self.titlestyle = Style(name="MyMaster-title", family="presentation")
        self.titlestyle.addElement(ParagraphProperties(textalign="center"))
        self.titlestyle.addElement(TextProperties(fontsize="30pt", fontfamily="sans"))
        self.titlestyle.addElement(
            GraphicProperties(fillcolor="#ffffff", strokecolor="#ffffff")
        )
        self.doc.styles.addElement(self.titlestyle)

        # Style for adding content
        self.teststyle = Style(name="MyMaster-subtitle", family="presentation")
        self.teststyle.addElement(ParagraphProperties(textalign="left"))
        self.teststyle.addElement(TextProperties(fontsize="14pt", fontfamily="sans"))
        self.teststyle.addElement(
            GraphicProperties(fillcolor="#ffffff", strokecolor="#ffffff")
        )
        self.doc.styles.addElement(self.teststyle)

        # Style for images
        # self.photostyle = Style(name="MyMaster-photo", family="presentation")
        # self.doc.styles.addElement(self.photostyle)

        # Style for pages and transitions
        self.dpstyle = Style(name="dp1", family="drawing-page")
        self.dpstyle.addElement(
            DrawingPageProperties(
                transitiontype="none",
                # transitionstyle="move-from-top",
                duration="PT00S",
            )
        )
        self.doc.automaticstyles.addElement(self.dpstyle)

        # Every drawing page must have a master page assigned to it.
        self.masterpage = MasterPage(name="MyMaster", pagelayoutname=pagelayout)
        self.doc.masterstyles.addElement(self.masterpage)

        # Metadata
        self.doc.meta.addElement(
            dc.Title(text="Shiftleader Report for week <week num>")
        )
        self.doc.meta.addElement(dc.Date(text=datetime.now().isoformat()))
        # self.doc.meta.addElement(dc.Subject(text="Shiftleader Report"))
        self.doc.meta.addElement(
            dc.Creator(text="CertHelper version <version>, requested by <username>")
        )

    def save(self, filename: str) -> None:
        self.doc.save(outputfile=filename)

    def add_page(self, title: str, content: str):
        page = Page(stylename=self.dpstyle, masterpagename=self.masterpage)
        self.doc.presentation.addElement(page)

        titleframe = Frame(
            stylename=self.titlestyle, width="25cm", height="2cm", x="1.5cm", y="2cm"
        )
        textbox = TextBox()
        titleframe.addElement(textbox)
        textbox.addElement(P(text=title))
        page.addElement(titleframe)

        if content:
            c_frame = Frame(
                stylename=self.teststyle,
                # stylename=self.titlestyle,
                width="25cm",
                height="4cm",
                x="1.5cm",
                y="7cm",
            )
            c_text = TextBox()
            c_frame.addElement(c_text)
            c_text.addElement(P(text=content))
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
