using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Apply a built‑in table style (any style will do).
            table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

            // No conditional style options are needed.
            table.StyleOptions = TableStyleOptions.None;

            // Hide all borders by clearing them directly (NoBorders enum does not exist).
            table.ClearBorders();

            // Prepare output directory.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "TableWithNoBorders.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Inform that the process completed.
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
