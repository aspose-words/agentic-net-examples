using System;
using System.Drawing;
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

            // Start a table.
            Table table = builder.StartTable();

            // Add a simple 2x2 table.
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

            // Apply a custom color to the top border of the table.
            // Use SetBorder to modify the top edge directly.
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.DarkCyan, true);

            // Save the document to the local file system.
            string outputPath = "TableWithCustomTopBorder.docx";
            doc.Save(outputPath);
        }
    }
}
