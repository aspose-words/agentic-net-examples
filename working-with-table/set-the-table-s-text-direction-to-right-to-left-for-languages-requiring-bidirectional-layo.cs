using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBidiExample
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
            builder.Write("Cell 1 (LTR text)");
            builder.InsertCell();
            builder.Write("Cell 2 (LTR text)");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3 (LTR text)");
            builder.InsertCell();
            builder.Write("Cell 4 (LTR text)");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Set the table to be right‑to‑left (bidirectional).
            table.Bidi = true;

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBidi.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
