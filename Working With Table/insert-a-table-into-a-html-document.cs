using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoHtml
{
    static void Main()
    {
        // Path to the source HTML file.
        const string inputHtmlPath = @"C:\Docs\source.html";

        // Path where the modified HTML will be saved.
        const string outputHtmlPath = @"C:\Docs\result.html";

        // Load the existing HTML document.
        Document doc = new Document(inputHtmlPath);

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document so the table is appended.
        builder.MoveToDocumentEnd();

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Optional: adjust table formatting (example: set borders).
        table.SetBorders(LineStyle.Single, 1.0, System.Drawing.Color.Black);

        // Save the modified document back to HTML format.
        doc.Save(outputHtmlPath, SaveFormat.Html);
    }
}
