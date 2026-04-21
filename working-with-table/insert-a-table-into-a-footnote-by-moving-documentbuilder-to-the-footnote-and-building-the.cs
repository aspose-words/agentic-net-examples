using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some main text.
        builder.Write("Sample text with a footnote");

        // Insert a footnote with empty text (we will fill it with a table).
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "");

        // Move the builder cursor to the first paragraph of the footnote.
        builder.MoveTo(footnote.FirstParagraph);

        // Build a 2x2 table inside the footnote.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FootnoteTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Indicate successful completion.
        Console.WriteLine("Document saved to " + outputPath);
    }
}
