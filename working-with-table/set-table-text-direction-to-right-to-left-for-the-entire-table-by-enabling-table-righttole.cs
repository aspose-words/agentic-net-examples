using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRightToLeftExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a couple of cells with sample text.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1 (left-to-right)");
            builder.InsertCell();
            builder.Write("Cell 2 (left-to-right)");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3 (left-to-right)");
            builder.InsertCell();
            builder.Write("Cell 4 (left-to-right)");
            builder.EndRow();

            // Finish building the table.
            builder.EndTable();

            // Enable right‑to‑left layout for the entire table.
            // The Bidi property controls the text direction of a table.
            table.Bidi = true;

            // Save the document to the file system.
            const string outputPath = "TableRightToLeft.docx";
            doc.Save(outputPath);

            // Optional: indicate that the file was created (no console input required).
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
