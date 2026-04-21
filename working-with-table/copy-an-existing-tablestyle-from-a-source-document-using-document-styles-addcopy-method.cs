using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleCopy
{
    public class Program
    {
        public static void Main()
        {
            // Create a source document and define a custom table style.
            Document srcDoc = new Document();

            // Add a new table style named "MyTableStyle".
            // Cast the returned Style to TableStyle to access table‑specific formatting members.
            TableStyle srcTableStyle = (TableStyle)srcDoc.Styles.Add(StyleType.Table, "MyTableStyle");

            // Configure visual properties of the style.
            srcTableStyle.Shading.BackgroundPatternColor = Color.LightYellow;
            srcTableStyle.Borders.Color = Color.DarkBlue;
            srcTableStyle.Borders.LineStyle = LineStyle.Single;
            srcTableStyle.CellSpacing = 5.0;
            srcTableStyle.BottomPadding = 10.0;
            srcTableStyle.TopPadding = 10.0;
            srcTableStyle.LeftPadding = 8.0;
            srcTableStyle.RightPadding = 8.0;

            // Save the source document (optional, just to demonstrate that the style exists in a file).
            string srcPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
            srcDoc.Save(srcPath);

            // Create a destination document where the style will be copied.
            Document dstDoc = new Document();

            // Copy the table style from the source document into the destination document.
            Style copiedStyle = dstDoc.Styles.AddCopy(srcTableStyle);
            // Optionally rename the copied style.
            copiedStyle.Name = "CopiedTableStyle";

            // Build a simple table in the destination document and apply the copied style.
            DocumentBuilder builder = new DocumentBuilder(dstDoc);
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
            builder.EndRow();

            builder.EndTable();

            // Apply the copied style to the newly created table.
            table.Style = copiedStyle;

            // Save the destination document.
            string dstPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
            dstDoc.Save(dstPath);

            // Simple verification that the style was copied and applied.
            if (dstDoc.Styles["CopiedTableStyle"] != null && table.StyleName == "CopiedTableStyle")
            {
                Console.WriteLine("Table style successfully copied and applied.");
            }
            else
            {
                throw new InvalidOperationException("Failed to copy or apply the table style.");
            }
        }
    }
}
