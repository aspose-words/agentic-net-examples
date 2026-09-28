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

        // Start a new table.
        builder.StartTable();

        // ----- Header row (will repeat on each page) -----
        // Insert header cells.
        builder.InsertCell();
        builder.Write("Header Column 1");
        builder.InsertCell();
        builder.Write("Header Column 2");
        builder.InsertCell();
        builder.Write("Header Column 3");
        // End the header row.
        builder.EndRow();

        // Insert multiple data rows to force pagination.
        for (int i = 1; i <= 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} - Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} - Col 2");
            builder.InsertCell();
            builder.Write($"Row {i} - Col 3");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set the first row as a heading row that repeats on each page.
        Row headerRow = table.FirstRow;
        if (headerRow != null)
        {
            headerRow.RowFormat.HeadingFormat = true;
        }
        else
        {
            throw new Exception("Failed to locate the header row.");
        }

        // Save the document.
        string outputPath = "TableWithRepeatingHeader.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"The output file '{outputPath}' was not created.");
        }
    }
}
