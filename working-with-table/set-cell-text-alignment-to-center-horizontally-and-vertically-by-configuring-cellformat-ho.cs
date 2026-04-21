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

        // Start building a table.
        Table table = builder.StartTable();

        // ----- First Row -----
        // First cell.
        builder.InsertCell();
        // Center vertically.
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        // Center horizontally by aligning the paragraph.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // ----- Second Row (optional, shows same alignment) -----
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 3");

        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 4");

        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
