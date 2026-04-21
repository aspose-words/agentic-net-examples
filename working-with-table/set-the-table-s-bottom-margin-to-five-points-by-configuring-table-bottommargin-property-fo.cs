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
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndTable();

        // Set the table's bottom padding to five points.
        table.BottomPadding = 5.0;

        // Save the document to the local file system.
        string outputPath = "TableBottomPadding.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
