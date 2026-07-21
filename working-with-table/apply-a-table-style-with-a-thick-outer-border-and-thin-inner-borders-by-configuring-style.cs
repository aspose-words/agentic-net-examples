using System;
using System.IO;
using System.Drawing;
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

            // Build a simple 3x3 table.
            Table table = builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Create a custom table style.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

            // Configure the style's default cell borders (inner borders) to be thin.
            tableStyle.Borders.LineStyle = LineStyle.Single;
            tableStyle.Borders.LineWidth = 0.5; // thin inner border
            tableStyle.Borders.Color = Color.Black;

            // Apply the style to the table.
            table.Style = tableStyle;

            // Set thick outer borders directly on the table.
            table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

            // Define output path and save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithStyle.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved successfully.");
        }
    }
}
