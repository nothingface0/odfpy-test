from odf.opendocument import OpenDocumentText, OpenDocumentPresentation
from odf.style import Style, TextProperties
from odf.text import H, P, Span


def main():
    filename = "test.odp"
    doc = OpenDocumentPresentation()
    t = doc.createTextNode(data="asdfasdf")
    doc.save(outputfile=filename)


if __name__ == "__main__":
    main()
