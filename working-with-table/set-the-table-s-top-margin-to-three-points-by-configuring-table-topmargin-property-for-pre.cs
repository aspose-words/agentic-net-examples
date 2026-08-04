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

        // Build a simple table with one cell.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Set the distance between the table top and surrounding text to 3 points.
        table.DistanceTop = 3.0;

        // Save the document.
        string outputPath = "TableTopMargin.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
