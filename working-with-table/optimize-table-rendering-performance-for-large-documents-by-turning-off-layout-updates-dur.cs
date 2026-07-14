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

        // Use DocumentBuilder to construct a large table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // Add 1000 rows with two cells each.
        for (int i = 0; i < 1000; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Force a layout pass after all modifications.
        doc.UpdatePageLayout();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OptimizedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Confirmation message.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
