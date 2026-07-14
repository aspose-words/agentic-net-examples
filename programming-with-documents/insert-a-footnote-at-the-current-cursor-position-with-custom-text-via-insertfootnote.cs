using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text that will be referenced by the footnote.
        builder.Write("This is a sentence with a footnote reference.");

        // Insert a footnote at the current cursor position with custom text.
        builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text.");

        // Determine an output file path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FootnoteExample.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);
    }
}
