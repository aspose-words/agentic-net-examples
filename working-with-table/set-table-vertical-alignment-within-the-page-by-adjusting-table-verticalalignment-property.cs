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
        // Vertical alignment of a table within the page cannot be set directly.
        // If needed, you could adjust the table's position using other layout properties,
        // but for this example we simply leave the table with the default alignment.

        // Save the document to a file.
        string outputPath = "TableVerticalAlignment.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
