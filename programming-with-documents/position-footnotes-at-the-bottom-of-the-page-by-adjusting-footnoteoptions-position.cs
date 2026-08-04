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

        // Add some text and a footnote.
        builder.Write("Hello world! ");
        builder.InsertFootnote(FootnoteType.Footnote, "This is a footnote.");

        // Position footnotes at the bottom of each page.
        doc.FootnoteOptions.Position = FootnotePosition.BottomOfPage;

        // Prepare an output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FootnoteBottomOfPage.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
