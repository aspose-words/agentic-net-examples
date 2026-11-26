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

        // Create a custom table style with a larger font size.
        Style tableStyle = doc.Styles.Add(StyleType.Table, "MyTableStyle");
        tableStyle.Font.Size = 16; // larger font size for the whole table

        // Build a table with a header row.
        builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply the custom style and mark the first row as a header row.
        table.Style = tableStyle;
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document.
        string outputPath = "TableWithHeader.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Confirmation.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
