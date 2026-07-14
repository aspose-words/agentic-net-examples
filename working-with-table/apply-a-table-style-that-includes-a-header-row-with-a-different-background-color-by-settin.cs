using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace TableStyleHeaderExample
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
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Create a custom table style.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "HeaderBlueStyle");

            // Set the background color of the first row (header) via conditional style.
            tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

            // Apply the style to the table.
            table.Style = tableStyle;

            // Enable the first‑row conditional formatting.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleHeaderRow.docx");
            doc.Save(outputPath);
        }
    }
}
