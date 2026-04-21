using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleExample
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

            // Insert the first cell of the header row.
            builder.InsertCell();
            builder.Write("Product");
            // Insert the second cell of the header row.
            builder.InsertCell();
            builder.Write("Quantity");
            // Finish the header row.
            builder.EndRow();

            // Add a few data rows.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("35");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Enable the FirstRow option so the header row receives the style's conditional formatting.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Optional: adjust column widths to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Prepare the output folder.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Save the document.
            string outputPath = Path.Combine(artifactsDir, "TableWithHeaderStyle.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
