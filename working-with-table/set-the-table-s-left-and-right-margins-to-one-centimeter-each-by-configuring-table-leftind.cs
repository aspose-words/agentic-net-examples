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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Set the left indent to 1 cm (1 cm = 28.35 points).
        table.LeftIndent = 28.35;

        // Aspose.Words does not expose a RightIndent property.
        // To achieve a right‑side margin effect, align the table to the right.
        table.Alignment = TableAlignment.Right;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableMargins.docx");
        doc.Save(outputPath);
    }
}
