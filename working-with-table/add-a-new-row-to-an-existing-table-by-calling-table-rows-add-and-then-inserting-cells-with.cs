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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build an initial 2x2 table using the DocumentBuilder.
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

            // ------------------------------------------------------------
            // Add a new row to the existing table using Table.Rows.Add.
            // ------------------------------------------------------------

            // Create a new Row instance that belongs to the same document.
            Row newRow = new Row(doc);

            // Create the first cell for the new row.
            Cell cell1 = new Cell(doc);
            // Ensure the cell contains at least one paragraph.
            cell1.AppendChild(new Paragraph(doc));
            // Add text to the paragraph.
            cell1.FirstParagraph.AppendChild(new Run(doc, "R3C1"));
            // Add the cell to the row.
            newRow.Cells.Add(cell1);

            // Create the second cell for the new row.
            Cell cell2 = new Cell(doc);
            cell2.AppendChild(new Paragraph(doc));
            cell2.FirstParagraph.AppendChild(new Run(doc, "R3C2"));
            newRow.Cells.Add(cell2);

            // Append the new row to the table.
            table.Rows.Add(newRow);

            // Save the document to verify the result.
            string outputPath = "AddedRowTable.docx";
            doc.Save(outputPath);

            // Optional: confirm that the file was created.
            if (System.IO.File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to '{outputPath}'.");
            }
            else
            {
                throw new Exception("Failed to save the document.");
            }
        }
    }
}
