using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsInsertColumnExample
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
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Insert a new column at index 1 (between the existing columns).
            int insertIndex = 1;
            foreach (Row row in table.Rows)
            {
                // Create a new cell with an empty paragraph.
                Cell newCell = new Cell(doc);
                newCell.AppendChild(new Paragraph(doc));

                // Insert the cell at the desired position.
                if (insertIndex >= row.Cells.Count)
                    row.AppendChild(newCell); // Append if index is at the end.
                else
                    row.InsertBefore(newCell, row.Cells[insertIndex]);
            }

            // Populate the newly inserted column with sample text.
            for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
            {
                Cell cell = table.Rows[rowIdx].Cells[insertIndex];
                cell.FirstParagraph.AppendChild(new Run(doc, $"New{rowIdx + 1}"));
            }

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertColumn.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output file was not created.");
        }
    }
}
