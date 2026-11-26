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

        // Start building a table.
        builder.StartTable();

        // Header row cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Retrieve the table and mark the first row as a header that repeats on each page.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row headerRow = table.Rows[0];
        // The RowFormat.HeadingFormat property makes the row repeat as a header.
        headerRow.RowFormat.HeadingFormat = true;

        // Add many data rows to ensure the table spans multiple pages.
        for (int i = 0; i < 100; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Col 1");
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Col 2");
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Col 3");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "TableHeader.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
