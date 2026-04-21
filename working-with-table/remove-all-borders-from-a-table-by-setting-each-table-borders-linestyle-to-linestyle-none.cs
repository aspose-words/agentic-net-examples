using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x1 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample text");
        builder.EndTable();

        // Remove all borders from the table.
        table.ClearBorders();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableNoBorders.docx");
        doc.Save(outputPath);
    }
}
