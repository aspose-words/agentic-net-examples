using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row - header cells.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Second row - data cells.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("30");
        builder.EndRow();

        // Third row - data cells.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("45");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Optionally apply auto fit to contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a PDF file.
        doc.Save("TableInPdf.pdf", SaveFormat.Pdf);
    }
}
