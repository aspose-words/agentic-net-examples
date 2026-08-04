using System;
using System.IO;
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

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();
            builder.EndTable();

            // Create a custom table style and assign it to the table.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
            customStyle.CellSpacing = 5;
            customStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            customStyle.Borders.Color = Color.Blue;
            customStyle.Borders.LineStyle = LineStyle.DotDash;
            table.Style = customStyle;

            // Iterate through all styles in the document.
            foreach (Style style in doc.Styles)
            {
                // Process only table styles.
                if (style.Type == StyleType.Table)
                {
                    TableStyle tblStyle = (TableStyle)style;

                    // Example modifications:
                    // Increase cell spacing.
                    tblStyle.CellSpacing += 2;

                    // Change shading to a light gray.
                    tblStyle.Shading.BackgroundPatternColor = Color.LightGray;

                    // Set borders to a solid single line.
                    tblStyle.Borders.Color = Color.DarkGray;
                    tblStyle.Borders.LineStyle = LineStyle.Single;
                }
            }

            // Convert any remaining style formatting to direct formatting.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedTableStyles.docx");
            doc.Save(outputPath);
        }
    }
}
