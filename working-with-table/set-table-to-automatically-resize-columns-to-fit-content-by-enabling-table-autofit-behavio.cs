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
        Table table = builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity (kg)");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Enable auto‑fit so columns resize to fit their contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AutoFitTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
