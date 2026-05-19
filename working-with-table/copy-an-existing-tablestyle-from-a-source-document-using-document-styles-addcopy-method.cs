using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

namespace AsposeWordsTableStyleCopy
{
    public class Program
    {
        public static void Main()
        {
            // Create a source document and define a custom table style.
            Document srcDoc = new Document();

            // Add a new table style named "MyTableStyle" and cast it to TableStyle
            // to access style‑specific formatting properties.
            TableStyle srcStyle = (TableStyle)srcDoc.Styles.Add(StyleType.Table, "MyTableStyle");

            // Configure visual properties of the style.
            srcStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            srcStyle.Borders.Color = Color.Blue;
            srcStyle.Borders.LineStyle = LineStyle.DotDash;
            srcStyle.VerticalAlignment = CellVerticalAlignment.Center;

            // Build a simple table in the source document and apply the custom style.
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            Table srcTable = srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write("Source Table");
            srcBuilder.EndRow();
            srcBuilder.EndTable();
            srcTable.Style = srcStyle;

            // Save the source document to a temporary file.
            string srcPath = Path.Combine(Path.GetTempPath(), "Source.docx");
            srcDoc.Save(srcPath);

            // Create a destination document where the style will be copied.
            Document dstDoc = new Document();

            // Copy the style from the source document into the destination document.
            // AddCopy returns a Style; cast it to TableStyle for further use.
            TableStyle copiedStyle = (TableStyle)dstDoc.Styles.AddCopy(srcStyle);
            // Optionally rename the copied style for clarity.
            copiedStyle.Name = "CopiedTableStyle";

            // Build a table in the destination document and apply the copied style.
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
            Table dstTable = dstBuilder.StartTable();
            dstBuilder.InsertCell();
            dstBuilder.Write("Destination Table");
            dstBuilder.EndRow();
            dstBuilder.EndTable();
            dstTable.Style = copiedStyle;

            // Save the destination document.
            string dstPath = Path.Combine(Path.GetTempPath(), "Destination.docx");
            dstDoc.Save(dstPath);

            // Verify that the style exists in the destination document.
            if (dstDoc.Styles["CopiedTableStyle"] == null)
                throw new InvalidOperationException("The table style was not copied correctly.");
        }
    }
}
