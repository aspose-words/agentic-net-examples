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

            // Build an initial 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            // Finish the table construction.
            builder.EndTable();

            // Insert a new row at the end of the existing table.
            // Create a new Row node and append it to the table.
            Row newRow = new Row(doc);
            table.AppendChild(newRow);

            // Determine the number of columns from the first row.
            int columnCount = table.FirstRow.Cells.Count;

            // Populate each cell in the new row with placeholder text.
            for (int i = 0; i < columnCount; i++)
            {
                // Create a new cell.
                Cell cell = new Cell(doc);
                // Each cell must contain at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                // Add placeholder text.
                cell.FirstParagraph.AppendChild(new Run(doc, $"Placeholder {i + 1}"));
                // Append the cell to the new row.
                newRow.AppendChild(cell);
            }

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
        }
    }
}
