using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TablePreferredWidthExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single row with three cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell #1");
        builder.InsertCell();
        builder.Write("Cell #2");
        builder.InsertCell();
        builder.Write("Cell #3");
        builder.EndTable();

        // Set the table's preferred width to 100% of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Define an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablePreferredWidth.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
