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

        // Build an initial 2x2 table.
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

        // Finish the table construction.
        builder.EndTable();

        // Create a new row that will be inserted.
        Row newRow = new Row(doc);

        // First cell of the new row.
        Cell newCell1 = new Cell(doc);
        newCell1.AppendChild(new Paragraph(doc));
        newCell1.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 1"));
        newRow.AppendChild(newCell1);

        // Second cell of the new row.
        Cell newCell2 = new Cell(doc);
        newCell2.AppendChild(new Paragraph(doc));
        newCell2.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 2"));
        newRow.AppendChild(newCell2);

        // Insert the new row at index 1 (between the original rows).
        // Use the Rows collection's Insert method.
        table.Rows.Insert(1, newRow);

        // Save the document to disk.
        string outputFile = "InsertRowExample.docx";
        doc.Save(outputFile);

        // Simple verification that the file was created.
        if (File.Exists(outputFile))
        {
            Console.WriteLine($"Document successfully saved: {Path.GetFullPath(outputFile)}");
        }
        else
        {
            throw new InvalidOperationException("Failed to save the document.");
        }
    }
}
