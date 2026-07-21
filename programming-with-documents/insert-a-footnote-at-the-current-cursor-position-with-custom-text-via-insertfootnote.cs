using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to the document.
        builder.Write("This is a sample sentence with a footnote.");

        // Insert a footnote at the current cursor position with custom text.
        builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text.");

        // Prepare the output directory and file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FootnoteExample.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);
    }
}
