using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
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

            // Row 1, Cell 1
            builder.InsertCell();
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            builder.Write("Row 1, Cell 1");

            // Row 1, Cell 2
            builder.InsertCell();
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Row 2, Cell 1
            builder.InsertCell();
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            builder.Write("Row 2, Cell 1");

            // Row 2, Cell 2
            builder.InsertCell();
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Verify that each cell has bottom vertical alignment.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    if (cell.CellFormat.VerticalAlignment != CellVerticalAlignment.Bottom)
                        throw new InvalidOperationException("Cell vertical alignment is not set to Bottom.");
                }
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBottomAlignment.docx");
            doc.Save(outputPath);

            // Ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("Failed to create the output document.", outputPath);
        }
    }
}
