using System;
using System.IO;
using System.Drawing;
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

            // Build a simple 2x2 table.
            builder.StartTable();

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
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Create a new table style (optional – can be used for other formatting).
            Style tableStyle = doc.Styles.Add(StyleType.Table, "LightGrayTableStyle");
            // Note: The Style object does not expose a Table property in the current API version,
            // so we apply shading directly to the cells instead of via the style.

            // Retrieve the created table.
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];
            table.Style = tableStyle; // Apply the style (even though it has no shading defined).

            // Apply light gray shading to all cells in the table.
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                cell.CellFormat.Shading.Texture = TextureIndex.TextureNone;
                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }

            // Save the document.
            string outputPath = "TableStyleExample.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");
        }
    }
}
