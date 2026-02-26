using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AddRowExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document
            Document doc = new Document();

            // Create a table and add it to the document body
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Ensure the table has at least one row and one cell so we can work with it
            table.EnsureMinimum();

            // Add a new row to the existing table
            Row newRow = new Row(doc);
            table.AppendChild(newRow);

            // Add a single cell to the new row
            Cell newCell = new Cell(doc);
            newRow.AppendChild(newCell);

            // Add a paragraph with some text inside the new cell
            Paragraph para = new Paragraph(doc);
            newCell.AppendChild(para);
            para.AppendChild(new Run(doc, "This is a newly added row."));

            // Save the document
            doc.Save("AddedRow.docx");
        }
    }
}
