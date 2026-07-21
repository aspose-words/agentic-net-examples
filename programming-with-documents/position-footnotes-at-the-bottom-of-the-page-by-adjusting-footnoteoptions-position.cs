using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a footnote.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Hello world! ");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote contents.");

        // Position footnotes at the bottom of each page.
        doc.FootnoteOptions.Position = FootnotePosition.BottomOfPage;

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "FootnoteBottomOfPage.docx");
        doc.Save(outputPath);
    }
}
