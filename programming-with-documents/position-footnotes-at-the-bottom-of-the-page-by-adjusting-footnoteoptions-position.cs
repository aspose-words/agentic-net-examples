using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Define a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text and a footnote.
        builder.Write("Hello world! ");
        builder.InsertFootnote(FootnoteType.Footnote, "This footnote will appear at the bottom of the page.");

        // Explicitly set footnotes to be positioned at the bottom of each page.
        doc.FootnoteOptions.Position = FootnotePosition.BottomOfPage;

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "FootnoteBottomOfPage.docx");
        doc.Save(outputPath);
    }
}
