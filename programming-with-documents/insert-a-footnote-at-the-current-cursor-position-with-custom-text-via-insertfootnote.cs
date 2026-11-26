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

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some sample text to the document.
        builder.Write("This is a sentence that will have a footnote.");

        // Insert a footnote at the current cursor position with custom text.
        builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text.");

        // Define the output file path (saved in the current working directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FootnoteExample.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);
    }
}
