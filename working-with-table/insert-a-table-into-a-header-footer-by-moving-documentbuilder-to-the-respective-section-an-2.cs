using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class HeaderFooterTableExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table into the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        BuildSampleTable(builder, "Header");

        // Insert a table into the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        BuildSampleTable(builder, "Footer");

        // Save the document.
        string outputPath = "HeaderFooterTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }

    // Helper method that builds a simple 2x2 table with a title.
    private static void BuildSampleTable(DocumentBuilder builder, string title)
    {
        // Write a title above the table.
        builder.Writeln($"{title} Table:");

        // Start the table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write($"{title} Cell 1");
        builder.InsertCell();
        builder.Write($"{title} Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write($"{title} Cell 3");
        builder.InsertCell();
        builder.Write($"{title} Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Add a paragraph after the table to ensure proper layout.
        builder.Writeln();
    }
}
