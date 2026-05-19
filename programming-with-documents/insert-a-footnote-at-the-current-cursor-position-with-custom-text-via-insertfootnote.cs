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

        // Write some text that will be referenced by the footnote.
        builder.Write("This is a sentence with a footnote.");

        // Insert a footnote at the current cursor position with custom text.
        builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text.");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FootnoteExample.docx");
        doc.Save(outputPath);
    }
}
