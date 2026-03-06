using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToRtf
{
    static void Main()
    {
        // Load the existing RTF document.
        Document doc = new Document("InputDocument.rtf");

        // Ensure the document contains at least one table.
        if (doc.FirstSection.Body.Tables.Count == 0)
            throw new InvalidOperationException("The document does not contain any tables.");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Ensure the row has at least one cell (required for a valid row).
        newRow.EnsureMinimum();

        // Optionally add content to the first cell of the new row.
        // Here we insert a paragraph with some text.
        Paragraph para = new Paragraph(doc);
        para.AppendChild(new Run(doc, "New row added programmatically."));
        // Add the paragraph to the first cell's paragraph collection (FirstParagraph is read‑only).
        newRow.FirstCell.Paragraphs.Add(para);

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Save the modified document back to RTF format.
        doc.Save("OutputDocument.rtf");
    }
}
