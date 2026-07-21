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

        // Configure horizontal alignment by setting paragraph alignment.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        // Configure vertical alignment for the cell.
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // First cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // Add a second row with the same alignment settings.
        builder.InsertCell();
        builder.Write("Cell 3");

        builder.InsertCell();
        builder.Write("Cell 4");

        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");
    }
}
