using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleDemo
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
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            // Create a custom table style named "CustomStyle".
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "CustomStyle");

            // Define shading for the style (light gray background).
            customStyle.Shading.BackgroundPatternColor = Color.LightGray;
            customStyle.Shading.Texture = TextureIndex.TextureSolid;

            // Define borders for the style (dark blue single line).
            customStyle.Borders.Color = Color.DarkBlue;
            customStyle.Borders.LineStyle = LineStyle.Single;
            // Optionally set border width for each side.
            customStyle.Borders.Left.LineWidth = 1.5;
            customStyle.Borders.Right.LineWidth = 1.5;
            customStyle.Borders.Top.LineWidth = 1.5;
            customStyle.Borders.Bottom.LineWidth = 1.5;

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Save the document to a file in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTableStyle.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // The program ends automatically; no user interaction required.
        }
    }
}
