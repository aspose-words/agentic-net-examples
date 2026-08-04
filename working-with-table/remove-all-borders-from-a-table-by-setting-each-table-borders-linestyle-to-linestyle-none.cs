using System;
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
        builder.Write("Sample text");
        builder.EndTable();

        // Remove all borders from the table (including cell borders).
        table.ClearBorders();

        // Save the document.
        string outputPath = "TableNoBorders.docx";
        doc.Save(outputPath);
    }
}
