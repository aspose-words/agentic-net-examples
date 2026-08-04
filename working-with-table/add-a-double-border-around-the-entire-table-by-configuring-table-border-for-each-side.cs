using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableBorderExample
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

            // Apply a double line border to each side of the table.
            // The last parameter (true) overrides any existing cell borders.
            table.SetBorder(BorderType.Left,   LineStyle.Double, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Right,  LineStyle.Double, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Top,    LineStyle.Double, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Double, 1.5, Color.Black, true);

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Table.DoubleBorder.docx");
            doc.Save(outputPath);
        }
    }
}
