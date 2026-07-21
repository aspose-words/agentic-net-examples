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

        // Build a simple table with three cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.EndRow();
        builder.EndTable();

        // Set the table's preferred width to 100 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Save the document to a local file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TablePreferredWidth.docx");
        doc.Save(outputPath);
    }
}
