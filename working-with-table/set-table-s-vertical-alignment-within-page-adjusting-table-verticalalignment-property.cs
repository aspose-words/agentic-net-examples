using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class TableVerticalAlignmentExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table and add a single cell with some text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Align the table vertically to the middle of the page.
        // For floating tables the RelativeVerticalAlignment property controls the vertical position.
        table.RelativeVerticalAlignment = VerticalAlignment.Center; // Middle alignment.

        // Save the document to disk.
        doc.Save("TableVerticalAlignment.docx");
    }
}
