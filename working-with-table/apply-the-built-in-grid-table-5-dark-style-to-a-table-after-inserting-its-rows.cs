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

            // Start a new table.
            Table table = builder.StartTable();

            // ---- First row (header) ----
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // ---- Second row ----
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // ---- Third row ----
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("35");
            builder.EndRow();

            // ---- Fourth row ----
            builder.InsertCell();
            builder.Write("Carrots");
            builder.InsertCell();
            builder.Write("50");
            builder.EndRow();

            // Apply the built‑in "Grid Table 5 Dark" style to the table.
            table.StyleIdentifier = StyleIdentifier.GridTable5Dark;

            // Optionally, let the table auto‑fit its contents.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Finish the table.
            builder.EndTable();

            // Prepare output folder and file path.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "GridTable5Dark.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
