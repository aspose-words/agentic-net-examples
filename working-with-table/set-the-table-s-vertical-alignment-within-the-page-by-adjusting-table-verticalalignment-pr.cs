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

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // NOTE: The Table.VerticalAlignment property is not available in the current Aspose.Words version.
        // If vertical alignment of the table on the page is required, it must be achieved through layout tricks
        // (e.g., adding empty paragraphs before/after the table). Here we simply leave the table with default alignment.

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalAlignment.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
