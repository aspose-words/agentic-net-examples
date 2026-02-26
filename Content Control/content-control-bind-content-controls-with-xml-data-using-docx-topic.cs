using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

class ContentControlBindingExample
{
    static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Add a custom XML part that will hold the data to bind.
        // -----------------------------------------------------------------
        string xmlContent = @"
            <books>
                <book>
                    <title>Everyday Italian</title>
                    <author>Giada De Laurentiis</author>
                </book>
                <book>
                    <title>The C Programming Language</title>
                    <author>Brian W. Kernighan, Dennis M. Ritchie</author>
                </book>
                <book>
                    <title>Learning XML</title>
                    <author>Erik T. Ray</author>
                </book>
            </books>";

        // The part ID must be a GUID string.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xmlContent);

        // -----------------------------------------------------------------
        // 2. Insert a plain‑text content control (structured document tag).
        // -----------------------------------------------------------------
        StructuredDocumentTag titleTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Book Title"
        };
        // Bind to the first <title> element.
        titleTag.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/title[1]", string.Empty);

        builder.Write("First book title: ");
        builder.InsertNode(titleTag);
        builder.Writeln();

        // -----------------------------------------------------------------
        // 3. Insert a repeating section to list all books.
        // -----------------------------------------------------------------
        // Create a table with a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Title");
        builder.InsertCell();
        builder.Write("Author");
        builder.EndRow();
        // Do NOT call EndTable yet – we need to add the repeating section row first.

        // Row that will contain the repeating section.
        Row repeatRow = new Row(doc);
        table.Rows.Add(repeatRow);

        // Repeating section content control at the row level.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
        // Bind the repeating section to the collection of <book> elements.
        repeatingSection.XmlMapping.SetMapping(xmlPart, "/books[1]/book", string.Empty);
        repeatRow.AppendChild(repeatingSection);

        // Inside the repeating section, add a repeating section item.
        StructuredDocumentTag repeatingItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
        repeatingSection.AppendChild(repeatingItem);

        // Row that will be repeated for each <book>.
        Row dataRow = new Row(doc);
        repeatingItem.AppendChild(dataRow);

        // Title cell bound to the current <book>'s <title>.
        StructuredDocumentTag cellTitle = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        cellTitle.XmlMapping.SetMapping(xmlPart, "./title[1]", string.Empty);
        dataRow.AppendChild(cellTitle);

        // Author cell bound to the current <book>'s <author>.
        StructuredDocumentTag cellAuthor = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        cellAuthor.XmlMapping.SetMapping(xmlPart, "./author[1]", string.Empty);
        dataRow.AppendChild(cellAuthor);

        // Finish the table.
        builder.EndTable();

        // -----------------------------------------------------------------
        // 4. Save the document.
        // -----------------------------------------------------------------
        doc.Save("ContentControlBinding.docx");
    }
}
