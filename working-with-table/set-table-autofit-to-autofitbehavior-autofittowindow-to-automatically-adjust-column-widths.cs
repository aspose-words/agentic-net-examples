using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AutoFitTable.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑column table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("First column");
        builder.InsertCell();
        builder.Write("Second column");
        builder.EndRow();
        builder.EndTable();

        // Apply AutoFitToWindow so column widths adjust to page margins.
        table.AutoFit(AutoFitBehavior.AutoFitToWindow);

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
