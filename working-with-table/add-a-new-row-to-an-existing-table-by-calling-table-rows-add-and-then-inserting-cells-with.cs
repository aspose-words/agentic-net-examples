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

            // Build an initial table with one row and two cells using DocumentBuilder.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Original Cell 1");
            builder.InsertCell();
            builder.Write("Original Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Retrieve the created table (the first table in the document body).
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a new row that will be added to the existing table.
            Row newRow = new Row(doc);

            // Add two new cells to the row.
            // Cell 1
            Cell cell1 = new Cell(doc);
            // Ensure the cell contains a paragraph before adding text.
            cell1.AppendChild(new Paragraph(doc));
            cell1.FirstParagraph.AppendChild(new Run(doc, "New Cell 1"));
            newRow.Cells.Add(cell1);

            // Cell 2
            Cell cell2 = new Cell(doc);
            cell2.AppendChild(new Paragraph(doc));
            cell2.FirstParagraph.AppendChild(new Run(doc, "New Cell 2"));
            newRow.Cells.Add(cell2);

            // Append the new row to the table.
            table.Rows.Add(newRow);

            // Optional validation: ensure the table now has two rows.
            if (table.Rows.Count != 2)
                throw new InvalidOperationException("The new row was not added correctly.");

            // Save the document to a file in the current directory.
            string outputPath = "TableWithAddedRow.docx";
            doc.Save(outputPath);
        }
    }
}
