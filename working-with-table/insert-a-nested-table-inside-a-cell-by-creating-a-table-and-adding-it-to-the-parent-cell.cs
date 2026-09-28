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

        // Build the outer table with two cells.
        builder.StartTable();

        // First cell of the outer table.
        builder.InsertCell();
        builder.Write("Outer Cell 1");

        // Second cell of the outer table – this will contain a nested table.
        builder.InsertCell();

        // Build the nested table inside the current cell.
        builder.StartTable();

        // First cell of the nested table.
        builder.InsertCell();
        builder.Write("Nested Cell 1");
        builder.EndRow();

        // Second cell of the nested table.
        builder.InsertCell();
        builder.Write("Nested Cell 2");
        builder.EndRow();

        // Finish the nested table.
        builder.EndTable();

        // Finish the outer table row and the outer table itself.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "NestedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
