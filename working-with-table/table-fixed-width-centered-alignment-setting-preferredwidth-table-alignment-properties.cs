using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin building a table.
        Table table = builder.StartTable();

        // Add a single cell with some sample text.
        builder.InsertCell();
        builder.Write("Centered Fixed Width Table");

        // Finish the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Set a fixed width for the table (e.g., 300 points).
        table.PreferredWidth = PreferredWidth.FromPoints(300);

        // Align the table to the center of the page.
        table.Alignment = TableAlignment.Center;

        // Save the document to disk.
        doc.Save("TableFixedWidthCentered.docx");
    }
}
