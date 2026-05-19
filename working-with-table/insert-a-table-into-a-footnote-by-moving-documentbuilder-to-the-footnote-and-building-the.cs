using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;   // Required for Footnote and FootnoteType

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text in the main body.
        builder.Write("This is a paragraph with a footnote reference.");

        // Insert a footnote with placeholder text.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote placeholder.");

        // Move the builder cursor into the footnote's paragraph.
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
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Return the cursor to the end of the main document.
        builder.MoveToDocumentEnd();

        // Save the document.
        string outputPath = "FootnoteTable.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");
    }
}
