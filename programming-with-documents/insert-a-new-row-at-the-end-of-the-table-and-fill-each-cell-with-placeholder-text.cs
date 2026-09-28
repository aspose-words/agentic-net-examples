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

        // Build a sample table with 2 rows and 3 columns.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Create a new row to be added at the end of the table.
        Row newRow = new Row(doc); // Use the document as the owner.

        // Fill each cell of the new row with placeholder text.
        int columnCount = table.Rows[0].Cells.Count;
        for (int i = 0; i < columnCount; i++)
        {
            // Create a new cell.
            Cell cell = new Cell(doc); // Use the document as the owner.

            // Add a paragraph and run with placeholder text.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, $"Placeholder {i + 1}");
            para.AppendChild(run);
            cell.AppendChild(para);

            // Add the cell to the new row.
            newRow.Cells.Add(cell);
        }

        // Append the new row to the table.
        table.Rows.Add(newRow);

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
