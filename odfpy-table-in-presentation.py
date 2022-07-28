from datetime import datetime
from odf.opendocument import OpenDocumentPresentation
from odf.style import (
    Style,
    MasterPage,
    PageLayout,
    PageLayoutProperties,
    TextProperties,
    GraphicProperties,
    ParagraphProperties,
    DrawingPageProperties,
    TableProperties,
)
from odf import dc
from odf.text import P, List, ListItem, ListLevelStyleBullet
from odf.presentation import Header
from odf.draw import Page, Frame, TextBox, Image
from odf.table import Table, TableColumn, TableRow, TableCell


def main():
    doc = OpenDocumentPresentation()

    # Page configuration
    pagelayout = PageLayout(name="MyLayout")
    doc.automaticstyles.addElement(pagelayout)
    pagelayout.addElement(
        PageLayoutProperties(
            margin="0cm",
            pagewidth="720pt",
            pageheight="540pt",
            printorientation="landscape",
        )
    )

    # Create a masterpage
    masterpage = MasterPage(name="Master1", pagelayoutname=pagelayout)
    doc.masterstyles.addElement(masterpage)

    masterpagecontent = MasterPage(name="Master2", pagelayoutname=pagelayout)
    doc.masterstyles.addElement(masterpagecontent)

    # Title style configuration
    # Style name must be in the format <master page name>-<element type>
    # If family="presentation", they can also contain ParagraphProperties
    # and TextProperties
    titlestyle = Style(name="Master1-text", family="presentation")
    titlestyle.addElement(
        ParagraphProperties(textalign="center", verticalalign="middle")
    )
    titlestyle.addElement(TextProperties(fontsize="44pt", fontfamily="sans"))
    titlestyle.addElement(GraphicProperties(fill="none", stroke="none"))
    doc.styles.addElement(titlestyle)

    # Add title page - Master1
    page = Page(masterpagename=masterpage)
    doc.presentation.addElement(page)
    titleframe = Frame(
        stylename=titlestyle,
        width="612pt",
        height="115.7pt",
        x="54pt",
        y="167.8pt",
    )
    textbox = TextBox()
    titleframe.addElement(textbox)

    textbox.addElement(P(text="Offline Shift Leader Report"))
    page.addElement(titleframe)

    #####################################
    doc.save("table_in_presentation.odp")


if __name__ == "__main__":
    main()
