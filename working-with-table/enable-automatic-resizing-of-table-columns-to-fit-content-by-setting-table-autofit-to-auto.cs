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

        // Build a simple 2x2 table with varying text lengths.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Short");
        builder.InsertCell();
        builder.Write("A much longer text that should cause the column to expand");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Another short");
        builder.InsertCell();
        builder.Write("Medium length");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Enable automatic column resizing to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAutoFit.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
