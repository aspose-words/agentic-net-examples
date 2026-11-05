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

        // Build an initial table with one row and two cells.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Add a new row to the table.
        Row newRow = new Row(doc);
        table.Rows.Add(newRow);

        // Add two cells to the new row.
        Cell cell1 = new Cell(doc);
        Cell cell2 = new Cell(doc);
        newRow.Cells.Add(cell1);
        newRow.Cells.Add(cell2);

        // Ensure each cell contains a paragraph before adding runs.
        cell1.AppendChild(new Paragraph(doc));
        cell2.AppendChild(new Paragraph(doc));

        // Insert text into the new cells.
        cell1.FirstParagraph?.AppendChild(new Run(doc, "New Row Cell 1"));
        cell2.FirstParagraph?.AppendChild(new Run(doc, "New Row Cell 2"));

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
