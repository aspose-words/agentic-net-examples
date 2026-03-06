using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Ensure the row has at least one cell (creates a cell with an empty paragraph).
        newRow.EnsureMinimum();

        // Add some text to the first cell of the new row.
        newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "New row text"));

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
