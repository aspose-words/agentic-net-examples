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

        // First row - first cell.
        builder.InsertCell();
        // Center-align the paragraph inside this cell.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 1, Row 1");

        // First row - second cell.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row - first cell.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 1, Row 2");

        // Second row - second cell.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The document was not saved correctly.");
        }
    }
}
