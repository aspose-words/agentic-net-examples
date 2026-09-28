using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a new table node.
        Table table = new Table(doc);

        // ----- First row -----
        Row row1 = new Row(doc);

        // First cell of first row.
        Cell cell11 = new Cell(doc);
        cell11.AppendChild(new Paragraph(doc));
        cell11.FirstParagraph.AppendChild(new Run(doc, "R1C1"));
        row1.AppendChild(cell11);

        // Second cell of first row.
        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "R1C2"));
        row1.AppendChild(cell12);

        // Add the first row to the table.
        table.AppendChild(row1);

        // ----- Second row -----
        Row row2 = new Row(doc);

        // First cell of second row.
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "R2C1"));
        row2.AppendChild(cell21);

        // Second cell of second row.
        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "R2C2"));
        row2.AppendChild(cell22);

        // Add the second row to the table.
        table.AppendChild(row2);

        // Append the table to the document body.
        doc.FirstSection.Body.AppendChild(table);

        // Save the document.
        string outputPath = "CreatedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("Failed to create the output document.");
        }
    }
}
