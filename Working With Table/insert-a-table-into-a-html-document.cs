using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoHtml
{
    static void Main()
    {
        // Load the existing HTML document.
        Document doc = new Document("input.html");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired location).
        builder.MoveToDocumentEnd();

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Optionally, apply AutoFit to adjust column widths.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document back to HTML format.
        doc.Save("output.html", SaveFormat.Html);
    }
}
