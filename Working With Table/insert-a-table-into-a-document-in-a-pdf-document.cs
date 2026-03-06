using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which simplifies inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

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

        // Optionally adjust table formatting (e.g., auto‑fit to contents).
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a PDF file.
        doc.Save("TableInPdf.pdf", SaveFormat.Pdf);
    }
}
