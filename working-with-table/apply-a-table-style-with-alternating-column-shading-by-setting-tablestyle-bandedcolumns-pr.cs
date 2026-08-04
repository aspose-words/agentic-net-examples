using System;
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

            // Start a table and add a few rows and columns.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.InsertCell();
            builder.Write("Header 3");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            builder.InsertCell();
            builder.Write("Row 1, Col 2");
            builder.InsertCell();
            builder.Write("Row 1, Col 3");
            builder.EndRow();

            // Third row.
            builder.InsertCell();
            builder.Write("Row 2, Col 1");
            builder.InsertCell();
            builder.Write("Row 2, Col 2");
            builder.InsertCell();
            builder.Write("Row 2, Col 3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply a built‑in style to the table.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Enable column banding (alternating column shading).
            table.StyleOptions = TableStyleOptions.ColumnBands;

            // Save the document.
            string outputPath = "TableWithColumnBanding.docx";
            doc.Save(outputPath);
        }
    }
}
