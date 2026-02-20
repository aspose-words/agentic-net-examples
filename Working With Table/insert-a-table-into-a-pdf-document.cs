using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class InsertTableIntoPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // First row - header cells.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Second row - data.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Third row - data.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optionally adjust table layout (auto‑fit to contents).
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a PDF file.
        string outputPath = "TableInPdf.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
