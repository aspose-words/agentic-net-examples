using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a new table and attach it to the document's first section body.
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // ------------------------------
            // First row with two cells.
            // ------------------------------

            // Create the first row.
            Row firstRow = new Row(doc);
            table.AppendChild(firstRow);

            // First cell of the first row.
            Cell cell11 = new Cell(doc);
            firstRow.AppendChild(cell11);
            Paragraph para11 = new Paragraph(doc);
            cell11.AppendChild(para11);
            para11.AppendChild(new Run(doc, "Row 1, Cell 1"));

            // Second cell of the first row.
            Cell cell12 = new Cell(doc);
            firstRow.AppendChild(cell12);
            Paragraph para12 = new Paragraph(doc);
            cell12.AppendChild(para12);
            para12.AppendChild(new Run(doc, "Row 1, Cell 2"));

            // ------------------------------
            // Second row with two cells.
            // ------------------------------

            Row secondRow = new Row(doc);
            table.AppendChild(secondRow);

            Cell cell21 = new Cell(doc);
            secondRow.AppendChild(cell21);
            Paragraph para21 = new Paragraph(doc);
            cell21.AppendChild(para21);
            para21.AppendChild(new Run(doc, "Row 2, Cell 1"));

            Cell cell22 = new Cell(doc);
            secondRow.AppendChild(cell22);
            Paragraph para22 = new Paragraph(doc);
            cell22.AppendChild(para22);
            para22.AppendChild(new Run(doc, "Row 2, Cell 2"));

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CreatedTable.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to create the output document.");
            }
        }
    }
}
