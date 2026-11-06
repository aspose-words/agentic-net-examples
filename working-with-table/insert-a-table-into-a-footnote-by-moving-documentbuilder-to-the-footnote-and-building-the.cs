using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;   // Needed for Footnote and FootnoteType

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("This paragraph contains a footnote reference.");

        // Insert a footnote and obtain its reference.
        Footnote footnote = builder.InsertFootnote(
            FootnoteType.Footnote,
            "Initial footnote text.");

        // Move the builder to the footnote's first paragraph so we can add content there.
        builder.MoveTo(footnote.FirstParagraph);

        // Build a 2x2 table inside the footnote.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "FootnoteTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Indicate successful completion.
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
    }
}
