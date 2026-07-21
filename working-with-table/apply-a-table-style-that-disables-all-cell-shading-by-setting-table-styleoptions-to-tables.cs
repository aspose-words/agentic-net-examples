using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleNoShadingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a simple 2x2 grid.
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

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

            // Disable all conditional style options (no first row, no banding, etc.).
            table.StyleOptions = TableStyleOptions.None;

            // Remove any shading that might be present on the table cells.
            table.ClearShading();

            // Save the document to the local file system.
            string outputPath = "TableStyleNoShading.docx";
            doc.Save(outputPath);
        }
    }
}
