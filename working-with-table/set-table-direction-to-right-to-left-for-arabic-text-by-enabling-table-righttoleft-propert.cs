using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableDirection
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑column table with Arabic and English text.
            Table table = builder.StartTable();

            // First cell – Arabic text.
            builder.InsertCell();
            builder.Write("مرحبا بالعالم"); // "Hello World" in Arabic.

            // Second cell – English text.
            builder.InsertCell();
            builder.Write("Hello World");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Enable right‑to‑left layout for the table.
            // The Bidi property makes the table render its cells from right to left.
            table.Bidi = true;

            // Save the document to a file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRightToLeft.docx");
            doc.Save(outputPath);

            // Simple validation: ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
