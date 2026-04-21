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

        // ---------- First Row ----------
        // Set paragraph spacing before and after the row.
        builder.ParagraphFormat.SpaceBefore = 5;   // 5 points before the row.
        builder.ParagraphFormat.SpaceAfter = 10;   // 10 points after the row.

        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        // End the first row.
        builder.EndRow();

        // Reset spacing for the next row.
        builder.ParagraphFormat.SpaceBefore = 8;
        builder.ParagraphFormat.SpaceAfter = 12;

        // ---------- Second Row ----------
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Reset spacing for the third row.
        builder.ParagraphFormat.SpaceBefore = 3;
        builder.ParagraphFormat.SpaceAfter = 6;

        // ---------- Third Row ----------
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RowSpacing.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Inform the user where the document was saved.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
