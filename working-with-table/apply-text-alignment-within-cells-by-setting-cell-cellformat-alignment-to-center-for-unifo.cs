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

        // Start a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        // Align paragraph text to the center inside the cell.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Row 1, Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Row 2, Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Row 2, Cell 2");
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
