using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build an initial table with two rows and two columns.
        builder.StartTable();
        // First row
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();
        // Second row
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Create a new row and add it to the table.
        Row newRow = new Row(doc);
        table.Rows.Add(newRow);

        // Create the first new cell, add it to the row, and set its text.
        Cell cell1 = new Cell(doc);
        newRow.Cells.Add(cell1);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "New Row Cell 1"));

        // Create the second new cell, add it to the row, and set its text.
        Cell cell2 = new Cell(doc);
        newRow.Cells.Add(cell2);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "New Row Cell 2"));

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Inform that the process completed.
        Console.WriteLine("Document saved successfully to " + Path.GetFullPath(outputPath));
    }
}
