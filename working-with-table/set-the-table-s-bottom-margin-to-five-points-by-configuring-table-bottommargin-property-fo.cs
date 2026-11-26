using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBottomMarginExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Writeln("Cell 1,1");

            // First row, second cell.
            builder.InsertCell();
            builder.Writeln("Cell 1,2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Writeln("Cell 2,1");

            // Second row, second cell.
            builder.InsertCell();
            builder.Writeln("Cell 2,2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Set the bottom padding of the table to five points.
            // Aspose.Words uses BottomPadding instead of BottomMargin for this purpose.
            table.BottomPadding = 5f;

            // Save the document to disk.
            string outputPath = "TableBottomMargin.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
