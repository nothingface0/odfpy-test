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
    TableColumnProperties,
    TableRowProperties,
)
from odf import dc
from odf.text import P, List, ListItem, ListLevelStyleBullet, PageNumber, H, Span
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

    dpstyle = Style(name="dp1", family="drawing-page")
    dpstyle.addElement(
        DrawingPageProperties(
            transitiontype="none",
            # transitionstyle="move-from-top",
            duration="PT00S",
            displaypagenumber="true",
            displayfooter="true",
        )
    )
    doc.automaticstyles.addElement(dpstyle)

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

    WIDTH_FRAME = 648
    HEIGHT_FRAME = 105

    frame = Frame(
        stylename=style_master1_table,
        width=f"{WIDTH_FRAME}pt",
        height=f"{HEIGHT_FRAME}pt",
        x="40pt",
        y="117pt",
    )
    page.addElement(frame)

    # Style for table
    style_table = Style(name="tab1", family="table")
    style_table.addElement(TableProperties(writingmode="lr-tb"))

    # Style for cells
    style_cell = Style(name="ce1", family="table-cell")
    style_cell.addElement(
        GraphicProperties(
            fillcolor="#dddddd",
            textareaverticalalign="middle",
        )
    )
    style_cell.addElement(ParagraphProperties(writingmode="lr-tb", textalign="right"))
    style_cell.addElement(
        TableCellProperties(
            paddingtop="1in",
            paddingleft="0.1in",
            paddingright="0.1in",
            backgroundcolor="#dddddd",
        )
    )
    doc.automaticstyles.addElement(style_cell)

    # Style for cells in header
    style_cell_header = Style(name="ce2", family="table-cell")
    style_cell_header.addElement(
        GraphicProperties(
            fillcolor="#aaaaaa",
            textareaverticalalign="middle",
        )
    )
    style_cell_header.addElement(
        ParagraphProperties(writingmode="lr-tb", textalign="right")
    )
    style_cell_header.addElement(
        TableCellProperties(
            paddingtop="0.05in",
            paddingleft="0.1in",
            paddingright="0.1in",
            backgroundcolor="#aaaaaa",
        )
    )
    doc.automaticstyles.addElement(style_cell_header)

    style_span = Style(name="sp1", family="text")
    style_span.addElement(TextProperties(fontfamily="sans"))
    doc.automaticstyles.addElement(style_span)

    NUM_ROWS = 10
    NUM_COLS = 5
    WIDTH_COL = 100
    HEIGHT_ROW = 50

    table = Table(usebandingcolumnsstyles="false", usebandingrowsstyles="false")

    # Table *must* amount to total frame width. Cols are autosized if total
    # width is less than frame width

    # Style for rows
    style_row = Style(name="roh", family="table-row")
    style_row.addElement(TableRowProperties(rowheight=f"{HEIGHT_ROW}pt"))
    doc.automaticstyles.addElement(style_row)
    tr = TableRow(defaultcellstylename=style_cell_header, stylename=style_row)

    for i in range(NUM_COLS):
        # Last col gets remaining width
        width_col = (
            WIDTH_COL if i < NUM_COLS - 1 else WIDTH_FRAME - (NUM_COLS - 1) * WIDTH_COL
        )
        # Style for columns
        style_column = Style(name=f"co{i}", family="table-column")
        style_column.addElement(TableColumnProperties(columnwidth=f"{width_col}pt"))
        doc.automaticstyles.addElement(style_column)

        table.addElement(TableColumn(defaultcellstylename="", stylename=style_column))
        tc = TableCell()
        p = P()
        p.addElement(Span(text=f"Header{i}", stylename=style_span))
        tc.addElement(p)
        tr.addElement(tc)
    table.addElement(tr)

    for i in range(NUM_ROWS):
        height_row = (
            HEIGHT_ROW
            if i < NUM_ROWS - 1
            else HEIGHT_FRAME - (NUM_ROWS - 1) * HEIGHT_ROW
        )
        style_row = Style(name=f"ro{i}", family="table-row")
        style_row.addElement(TableRowProperties(rowheight=f"{height_row}pt"))
        doc.automaticstyles.addElement(style_row)
        tr = TableRow(defaultcellstylename=style_cell, stylename=style_row)

        for j in range(NUM_COLS):
            tc = TableCell()
            p = P()
            p.addElement(Span(text=f"cell({i:2d},{j:2d})", stylename=style_span))
            tc.addElement(p)
            tr.addElement(tc)
        table.addElement(tr)

    frame.addElement(table)

    #####################################
    doc.save("table_in_presentation.odp")


if __name__ == "__main__":
    main()
