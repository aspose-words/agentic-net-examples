using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with three rows and two columns.
        builder.StartTable();

        // Row 0
        builder.InsertCell();
        builder.Write("Row 0, Cell 0");
        builder.InsertCell();
        builder.Write("Row 0, Cell 1");
        builder.EndRow();

        // Row 1
        builder.InsertCell();
        builder.Write("Row 1, Cell 0");
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("Row 2, Cell 0");
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Delete the second row (index 1) using Table.Rows[index].Remove().
        if (table.Rows.Count > 1)
        {
            table.Rows[1].Remove();
        }
        else
        {
            throw new InvalidOperationException("The table does not contain enough rows to delete the specified row.");
        }

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
        {
            throw new IOException("The document was not saved correctly.");
        }
    }
}
