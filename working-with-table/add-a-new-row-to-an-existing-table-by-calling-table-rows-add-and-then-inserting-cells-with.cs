using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and build an initial 2x2 table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.StartTable();
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row to be added.
        Row newRow = new Row(doc);

        // First cell of the new row.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "R3C1"));
        newRow.Cells.Add(cell1);

        // Second cell of the new row.
        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "R3C2"));
        newRow.Cells.Add(cell2);

        // Add the new row to the table.
        table.Rows.Add(newRow);

        // Save the document with the added row.
        doc.Save("AddedRowTable.docx");
    }
}
