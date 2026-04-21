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

            // Start building a table.
            Table table = builder.StartTable();

            // ----- First row (header) -----
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Quantity");

            // End the header row.
            builder.EndRow();

            // ----- Second row -----
            builder.InsertCell();
            builder.Writeln("Apples");
            builder.InsertCell();
            builder.Writeln("20");
            builder.EndRow();

            // ----- Third row -----
            builder.InsertCell();
            builder.Writeln("Bananas");
            builder.InsertCell();
            builder.Writeln("40");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

            // Enable the FirstRow style option so that conditional formatting for the first row is applied.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // The built‑in style does not make the first row bold, so set the font weight manually.
            foreach (Cell cell in table.FirstRow.Cells)
            {
                // Ensure the cell contains at least one paragraph and one run.
                if (cell.FirstParagraph?.Runs.Count > 0)
                {
                    cell.FirstParagraph.Runs[0].Font.Bold = true;
                }
            }

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleFirstRowBold.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
