using System;
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

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build an initial 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndRow();

            // Finish the table construction.
            builder.EndTable();

            // Insert a new row at the end of the existing table.
            Row newRow = new Row(doc);
            // Assuming the table has two columns, create two cells with placeholder text.
            for (int col = 1; col <= 2; col++)
            {
                Cell cell = new Cell(doc);
                // Each cell must contain at least one paragraph.
                Paragraph para = new Paragraph(doc);
                cell.AppendChild(para);
                // Add placeholder text to the paragraph.
                Run run = new Run(doc, $"Placeholder {col}");
                para.AppendChild(run);
                // Append the cell to the new row.
                newRow.AppendChild(cell);
            }

            // Append the new row to the table (adds it as the last row).
            table.AppendChild(newRow);

            // Save the document to the local file system.
            doc.Save("Output.docx");
        }
    }
}
