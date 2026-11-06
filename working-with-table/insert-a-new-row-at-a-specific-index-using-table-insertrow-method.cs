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

        // Build a simple table with two rows and two columns.
        builder.StartTable();

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

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Clone the first row to use as a template for the new row.
        Row newRow = (Row)table.Rows[0].Clone(true);

        // Change the text in each cell of the new row.
        foreach (Cell cell in newRow.Cells)
        {
            // Ensure the cell has at least one run to modify.
            if (cell.FirstParagraph?.Runs.Count > 0)
                cell.FirstParagraph.Runs[0].Text = "Inserted";
            else
                cell.FirstParagraph?.AppendChild(new Run(doc, "Inserted"));
        }

        // Insert the new row at index 1 (between the original rows).
        table.Rows.Insert(1, newRow);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
