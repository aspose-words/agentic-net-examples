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

        // Use DocumentBuilder to construct the table (preferred workflow).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // ----- First row -----
        // First cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");

        // End the first row.
        builder.EndRow();

        // ----- Second row -----
        // First cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "CreatedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
