using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

namespace AsposeWordsTableStyleUpdate
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Add a custom table style to the document.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
            customStyle.CellSpacing = 5;
            customStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            customStyle.Borders.Color = Color.Blue;
            customStyle.Borders.LineStyle = LineStyle.DotDash;
            customStyle.VerticalAlignment = CellVerticalAlignment.Center;

            // Apply the custom style to the created table.
            table.Style = customStyle;

            // Iterate through all styles in the document and modify table style definitions.
            foreach (Style style in doc.Styles)
            {
                if (style.Type == StyleType.Table)
                {
                    TableStyle tableStyle = (TableStyle)style;

                    // Example modifications: change cell spacing, shading color, and border appearance.
                    tableStyle.CellSpacing = 10; // Increase spacing between cells.
                    tableStyle.Shading.BackgroundPatternColor = Color.LightGray; // New background color.
                    tableStyle.Borders.Color = Color.DarkRed; // New border color.
                    tableStyle.Borders.LineStyle = LineStyle.Single; // Solid border line.
                }
            }

            // Convert any style-based formatting to direct formatting so the changes are reflected in the table.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedTableStyle.docx");
            doc.Save(outputPath);
        }
    }
}
