using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleUpdate
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑cell table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Create a custom table style and set some initial formatting.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");
            customStyle.CellSpacing = 5;
            customStyle.Shading.BackgroundPatternColor = Color.LightGray;
            customStyle.Borders.Color = Color.Blue;
            customStyle.Borders.LineStyle = LineStyle.Single;

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Iterate through all styles in the document.
            foreach (Style style in doc.Styles)
            {
                // Process only table styles.
                if (style.Type == StyleType.Table)
                {
                    TableStyle tableStyle = (TableStyle)style;

                    // Update style properties programmatically.
                    tableStyle.CellSpacing = 10; // Increase spacing between cells.
                    tableStyle.Shading.BackgroundPatternColor = Color.Yellow; // Change background color.
                    tableStyle.Borders.LineStyle = LineStyle.DotDash; // Change border line style.
                    tableStyle.Borders.Color = Color.DarkGreen; // Change border color.
                }
            }

            // Save the document with the updated table style.
            doc.Save("UpdatedTableStyle.docx");
        }
    }
}
