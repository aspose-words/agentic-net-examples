using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑column table with a header row.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            builder.InsertCell();
            builder.Write("Row 1, Col 2");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Row 2, Col 1");
            builder.InsertCell();
            builder.Write("Row 2, Col 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");

            // Set the background color for the header (first) row.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightGray;

            // Optional: set a default background for the rest of the table.
            customStyle.Shading.BackgroundPatternColor = Color.White;

            // Apply the style to the table.
            table.Style = customStyle;

            // Enable the first‑row conditional formatting.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableStyleHeaderRow.docx");
            doc.Save(outputPath);
        }
    }
}
