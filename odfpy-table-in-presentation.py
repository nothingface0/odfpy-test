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
    TableCellProperties,
)
from odf import dc
from odf.text import P, List, ListItem, ListLevelStyleBullet, PageNumber, H
from odf.presentation import Header
from odf.draw import Page, Frame, TextBox, Image
from odf.table import (
    Table,
    TableColumn,
    TableRow,
    TableCell,
    TableHeaderRows,
    TableHeaderColumns,
)


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
    style_master1_title = Style(name="Master1-title", family="presentation")
    style_master1_title.addElement(
        ParagraphProperties(textalign="center", verticalalign="middle")
    )
    style_master1_title.addElement(TextProperties(fontsize="44pt", fontfamily="sans"))
    style_master1_title.addElement(GraphicProperties(fill="none", stroke="none"))
    doc.styles.addElement(style_master1_title)

    # Try a style for tables
    style_master1_table = Style(name="Master1-table", family="presentation")
    style_master1_table.addElement(TextProperties(fontsize="20pt", fontfamily="Impact"))
    doc.styles.addElement(style_master1_table)

    # Add title page - Master1
    page = Page(masterpagename=masterpage)
    doc.presentation.addElement(page)
    titleframe = Frame(
        stylename=style_master1_title,
        width="612pt",
        height="115.7pt",
        x="54pt",
        y="167.8pt",
    )
    textbox = TextBox()
    titleframe.addElement(textbox)
    textbox.addElement(P(text="Offline Shift Leader Report"))
    page.addElement(titleframe)

    # Create a table on a new page
    page = Page(masterpagename=masterpage)
    doc.presentation.addElement(page)
    frame = Frame(
        width="648pt",
        height="105pt",
        x="40pt",
        y="117pt",
    )
    page.addElement(frame)

    prop_graphic = GraphicProperties(fillcolor="#ffffff")
    prop_par = ParagraphProperties(textalign="center")

    ## Style for cells
    style_cell = Style(name="ce1", family="table-cell")
    style_cell.addElement(prop_graphic)
    doc.automaticstyles.addElement(style_cell)

    table = Table(name="Table1", usefirstrowstyles="true")
    table.addElement(TableColumn(numbercolumnsrepeated=2))

    tr = TableRow()
    tc = TableCell()
    tc.addElement(P(text="Header1"))
    tr.addElement(tc)
    tc = TableCell()
    tc.addElement(P(text="Header2"))
    tr.addElement(tc)
    table.addElement(tr)

    tr = TableRow(defaultcellstylename=style_cell)
    tc = TableCell()
    tc.addElement(P(text="cell00"))
    tr.addElement(tc)
    tc = TableCell()
    tc.addElement(P(text="cell01"))
    tr.addElement(tc)
    table.addElement(tr)

    frame.addElement(table)

    #####################################
    doc.save("table_in_presentation.odp")


if __name__ == "__main__":
    main()
