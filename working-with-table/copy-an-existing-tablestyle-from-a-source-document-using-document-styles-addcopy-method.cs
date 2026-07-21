using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

namespace TableStyleCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Define output folder (current directory)
            string outputDir = Directory.GetCurrentDirectory();

            // -------------------------
            // Create source document with a custom table style
            // -------------------------
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

            // Build a simple 2x2 table
            Table srcTable = srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write("Cell 1");
            srcBuilder.InsertCell();
            srcBuilder.Write("Cell 2");
            srcBuilder.EndRow();

            srcBuilder.InsertCell();
            srcBuilder.Write("Cell 3");
            srcBuilder.InsertCell();
            srcBuilder.Write("Cell 4");
            srcBuilder.EndTable();

            // Create a custom table style in the source document
            TableStyle srcTableStyle = (TableStyle)srcDoc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
            srcTableStyle.AllowBreakAcrossPages = true;
            srcTableStyle.CellSpacing = 5;
            srcTableStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            srcTableStyle.Borders.Color = Color.Blue;
            srcTableStyle.Borders.LineStyle = LineStyle.DotDash;
            srcTableStyle.VerticalAlignment = CellVerticalAlignment.Center;

            // Apply the custom style to the source table
            srcTable.Style = srcTableStyle;

            // Save the source document
            string srcPath = Path.Combine(outputDir, "Source.docx");
            srcDoc.Save(srcPath);

            // -------------------------
            // Create destination document and copy the table style from the source
            // -------------------------
            Document dstDoc = new Document();
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

            // Copy the style from the source document into the destination document
            // The AddCopy method creates a copy and automatically generates a new name if needed.
            Style copiedStyle = dstDoc.Styles.AddCopy(srcTableStyle);
            // Optionally rename the copied style for clarity
            copiedStyle.Name = "CopiedTableStyle";

            // Build another table in the destination document
            Table dstTable = dstBuilder.StartTable();
            dstBuilder.InsertCell();
            dstBuilder.Write("A");
            dstBuilder.InsertCell();
            dstBuilder.Write("B");
            dstBuilder.EndRow();

            dstBuilder.InsertCell();
            dstBuilder.Write("C");
            dstBuilder.InsertCell();
            dstBuilder.Write("D");
            dstBuilder.EndTable();

            // Apply the copied style to the new table
            dstTable.Style = copiedStyle;

            // Save the destination document
            string dstPath = Path.Combine(outputDir, "Destination.docx");
            dstDoc.Save(dstPath);

            // Simple verification that the style was copied
            if (dstDoc.Styles["CopiedTableStyle"] == null)
                throw new InvalidOperationException("The table style was not copied correctly.");

            // Program ends without waiting for user input
        }
    }
}
