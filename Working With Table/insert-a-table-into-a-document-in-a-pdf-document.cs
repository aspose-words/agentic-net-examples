using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which provides a convenient API for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder returns the Table node that was created.
        Table table = builder.StartTable();

        // ---- First row ----
        // Insert first cell and add text.
        builder.InsertCell();
        builder.Write("Cell 1,1");

        // Insert second cell in the same row and add text.
        builder.InsertCell();
        builder.Write("Cell 1,2");

        // End the first row.
        builder.EndRow();

        // ---- Second row ----
        builder.InsertCell();
        builder.Write("Cell 2,1");

        builder.InsertCell();
        builder.Write("Cell 2,2");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a PDF file.
        doc.Save("TableInPdf.pdf", SaveFormat.Pdf);
    }
}
