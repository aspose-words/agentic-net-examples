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

            // Build an initial 2‑column table with one row using the builder.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Original Row, Cell 1");
            builder.InsertCell();
            builder.Write("Original Row, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Retrieve the created table.
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a new row that will be added to the existing table.
            Row newRow = new Row(doc);
            // Add the new row to the table's Rows collection.
            table.Rows.Add(newRow);

            // Create the first cell for the new row.
            Cell newCell1 = new Cell(doc);
            // Ensure the cell contains a paragraph (required for text).
            newCell1.AppendChild(new Paragraph(doc));
            // Add text to the paragraph.
            newCell1.FirstParagraph.AppendChild(new Run(doc, "New Row, Cell 1"));
            // Add the cell to the row's Cells collection.
            newRow.Cells.Add(newCell1);

            // Create the second cell for the new row.
            Cell newCell2 = new Cell(doc);
            newCell2.AppendChild(new Paragraph(doc));
            newCell2.FirstParagraph.AppendChild(new Run(doc, "New Row, Cell 2"));
            newRow.Cells.Add(newCell2);

            // Save the document to verify the new row has been added.
            doc.Save("AddedRowTable.docx");
        }
    }
}
