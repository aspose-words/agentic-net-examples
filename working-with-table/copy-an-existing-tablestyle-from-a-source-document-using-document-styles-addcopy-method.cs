using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a source document and add a custom table style.
            Document srcDoc = new Document();
            // Add a new table style named "MyTableStyle".
            TableStyle srcTableStyle = (TableStyle)srcDoc.Styles.Add(StyleType.Table, "MyTableStyle");
            // Configure some visual properties of the style.
            srcTableStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            srcTableStyle.Borders.Color = Color.Blue;
            srcTableStyle.Borders.LineStyle = LineStyle.DotDash;
            srcTableStyle.CellSpacing = 5;
            srcTableStyle.VerticalAlignment = CellVerticalAlignment.Center;

            // (Optional) Save the source document for inspection.
            srcDoc.Save("Source.docx");

            // Create a destination document.
            Document dstDoc = new Document();

            // Copy the table style from the source document into the destination document.
            Style copiedStyle = dstDoc.Styles.AddCopy(srcDoc.Styles["MyTableStyle"]);

            // Build a simple table in the destination document and apply the copied style.
            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Assign the copied style to the new table.
            table.Style = copiedStyle;

            // Save the destination document.
            dstDoc.Save("Destination.docx");
        }
    }
}
