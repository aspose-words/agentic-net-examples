using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableDoubleBorderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a few cells with sample text.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Remove any existing borders.
            table.ClearBorders();

            // Apply a double line border to each side of the table.
            table.SetBorder(BorderType.Left,   LineStyle.Double, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Right,  LineStyle.Double, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Top,    LineStyle.Double, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Double, 1.5, Color.Black, true);

            // Save the document to a file.
            doc.Save("TableDoubleBorder.docx");
        }
    }
}
