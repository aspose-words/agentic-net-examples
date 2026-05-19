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
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

            // Configure the style's cell borders (inner borders) to be thin.
            customStyle.Borders.LineStyle = LineStyle.Single;
            customStyle.Borders.LineWidth = 0.5; // Thin inner borders.
            customStyle.Borders.Color = Color.Black;

            // Apply the style to the table.
            table.Style = customStyle;

            // Set thick outer borders directly on the table.
            table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithStyle.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
