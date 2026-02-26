using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Get the first table in the document (assumes at least one table exists).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Ensure the row has at least one cell with a paragraph.
        newRow.EnsureMinimum();

        // Add some text to the first cell of the new row.
        newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "This is a newly added row."));

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
