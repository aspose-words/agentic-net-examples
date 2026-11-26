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

        // Add a paragraph in the main body.
        builder.Writeln("This paragraph contains a footnote reference.");

        // Insert a footnote and obtain the Footnote node.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Initial footnote text.");

        // Move the builder to the first paragraph of the footnote so we can add content there.
        // Footnote is a CompositeNode; its first paragraph can be accessed via the Paragraphs collection.
        builder.MoveTo(footnote.FirstParagraph);

        // Build a 2x2 table inside the footnote.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "FootnoteTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful creation.
        Console.WriteLine("Document created: " + Path.GetFullPath(outputPath));
    }
}
