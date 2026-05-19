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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text and insert a footnote.
        builder.Write("Hello world! ");
        builder.InsertFootnote(FootnoteType.Footnote, "This is a footnote placed at the bottom of the page.");

        // Set footnote position to the bottom of each page.
        doc.FootnoteOptions.Position = FootnotePosition.BottomOfPage;

        // Prepare output directory and file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FootnoteBottomOfPage.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
