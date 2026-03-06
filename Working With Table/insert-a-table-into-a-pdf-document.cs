using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which provides a convenient API for building the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The method returns the Table node that was created.
        Table table = builder.StartTable();

        // ---- First row (header) ----
        builder.InsertCell();                     // First cell of the row.
        builder.Write("Header 1");                // Insert text into the first cell.
        builder.InsertCell();                     // Second cell of the row.
        builder.Write("Header 2");                // Insert text into the second cell.
        builder.EndRow();                         // Finish the first row.

        // ---- Second row (data) ----
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();                         // Finish the second row.

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in table style (optional).
        table.StyleIdentifier = StyleIdentifier.LightListAccent1;

        // Auto‑fit the table to its contents so the columns adjust to the text width.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a PDF file.
        doc.Save("TableInPdf.pdf", SaveFormat.Pdf);
    }
}
