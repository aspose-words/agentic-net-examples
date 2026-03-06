using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder returns the created Table node.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();               // First cell of the first row.
        builder.Write("Cell 1,1");           // Insert text into the cell.

        builder.InsertCell();               // Second cell of the first row.
        builder.Write("Cell 1,2");

        builder.EndRow();                   // End the first row.

        // ---- Second row ----
        builder.InsertCell();               // First cell of the second row.
        builder.Write("Cell 2,1");

        builder.InsertCell();               // Second cell of the second row.
        builder.Write("Cell 2,2");

        builder.EndRow();                   // End the second row.

        // Finish the table.
        builder.EndTable();

        // Save the document as a DOCX file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableExample.docx");
        doc.Save(outputPath);
    }
}
