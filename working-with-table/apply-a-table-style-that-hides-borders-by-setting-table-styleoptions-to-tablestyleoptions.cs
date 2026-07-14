using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleNoBordersExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑column table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Apply any built‑in style (optional, just to demonstrate style usage).
            table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

            // Hide all borders by clearing them directly (NoBorders enum does not exist).
            table.ClearBorders();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableNoBorders.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
