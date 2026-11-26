using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableVerticalAlignmentExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Writeln("Cell 3");
            builder.InsertCell();
            builder.Writeln("Cell 4");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Set vertical alignment of every cell to Bottom.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
                }
            }

            // Save the document.
            string outputPath = "VerticalAlignmentTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
