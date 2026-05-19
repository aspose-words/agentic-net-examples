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
        // Center vertically.
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        // Center horizontally by aligning the paragraph.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 4");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
