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

        // First cell: Arabic text, set paragraph direction to right‑to‑left.
        Cell arabicCell = builder.InsertCell();
        builder.ParagraphFormat.Bidi = true; // Apply RTL to the paragraph inside this cell.
        builder.Write("مرحبا بالعالم"); // "Hello World" in Arabic.

        // Second cell: English text, keep left‑to‑right direction.
        Cell englishCell = builder.InsertCell();
        builder.ParagraphFormat.Bidi = false; // Apply LTR to the paragraph inside this cell.
        builder.Write("Hello World");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellDirection.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Reload the document and confirm the paragraph direction of the Arabic cell.
        Document loadedDoc = new Document(outputPath);
        Cell loadedArabicCell = loadedDoc.FirstSection.Body.Tables[0].FirstRow.FirstCell;
        bool isRightToLeft = loadedArabicCell.FirstParagraph.ParagraphFormat.Bidi;
        if (!isRightToLeft)
            throw new Exception("The paragraph direction was not set to right‑to‑left.");

        // Execution completed successfully.
    }
}
