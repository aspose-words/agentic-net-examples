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
        builder.Write("Cell content");
        builder.EndTable();

        // Set the spacing before and after the table (in points).
        // Use DistanceTop and DistanceBottom to control the space between the table and surrounding text.
        table.DistanceTop = 12;
        table.DistanceBottom = 12;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableSpace.docx");
        doc.Save(outputPath);
    }
}
