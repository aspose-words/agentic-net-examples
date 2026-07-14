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

            // Use DocumentBuilder to construct a simple 2‑cell table.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();

            builder.InsertCell();
            builder.Write("Cell 1");

            builder.InsertCell();
            builder.Write("Cell 2");

            builder.EndRow();
            builder.EndTable();

            // Set the table to right‑to‑left (bidirectional) layout.
            table.Bidi = true;

            // Define an output path relative to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableBidi.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");

            // The program ends automatically; no user interaction required.
        }
    }
}
