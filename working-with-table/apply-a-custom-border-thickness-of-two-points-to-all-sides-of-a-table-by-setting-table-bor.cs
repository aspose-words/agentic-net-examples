using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableBorders
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Insert a few cells with sample text.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Apply a 2‑point single line border to all sides of the table.
            // This sets line style, width (in points) and color for every border.
            table.SetBorders(LineStyle.Single, 2.0, Color.Black);

            // Save the document to the local file system.
            string outputPath = "TableWithCustomBorders.docx";
            doc.Save(outputPath);
        }
    }
}
