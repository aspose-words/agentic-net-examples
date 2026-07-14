using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleShadingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and keep a reference to it.
            Table table = builder.StartTable();

            // First row with two cells.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row with two cells.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Create a custom table style.
            TableStyle grayStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyGrayStyle");

            // Apply a light gray background to all cells via the style's shading.
            grayStyle.Shading.BackgroundPatternColor = Color.LightGray;

            // Assign the custom style to the table.
            table.Style = grayStyle;

            // Save the document to the local file system.
            doc.Save("TableStyleShading.docx");
        }
    }
}
