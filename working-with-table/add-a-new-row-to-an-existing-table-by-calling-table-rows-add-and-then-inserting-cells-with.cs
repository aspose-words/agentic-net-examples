using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build an initial table with two rows and two cells per row.
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
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Create a new row.
        Row newRow = new Row(doc);

        // First cell.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "R3C1"));
        newRow.Cells.Add(cell1);

        // Second cell.
        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "R3C2"));
        newRow.Cells.Add(cell2);

        // Add the new row to the table.
        table.Rows.Add(newRow);

        // Save the document.
        string outputPath = "AddedRow.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
