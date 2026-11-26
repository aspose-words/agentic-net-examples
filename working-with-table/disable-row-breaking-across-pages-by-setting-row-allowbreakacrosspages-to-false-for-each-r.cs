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

        // Build a simple table with 5 rows and 3 columns.
        builder.StartTable();
        for (int i = 0; i < 5; i++)
        {
            for (int j = 0; j < 3; j++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {i + 1}, Cell {j + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Disable row breaking across pages for every row in the document.
        foreach (Row row in doc.GetChildNodes(NodeType.Row, true))
        {
            // The property is on RowFormat, not directly on Row.
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");
    }
}
