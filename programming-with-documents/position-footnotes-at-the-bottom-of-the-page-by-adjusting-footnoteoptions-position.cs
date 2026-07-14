using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text and add a footnote.
        builder.Write("Hello world! ");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote contents.");

        // Position all footnotes at the bottom of each page.
        doc.FootnoteOptions.Position = FootnotePosition.BottomOfPage;

        // Ensure the output folder exists.
        string artifactsDir = "Artifacts";
        System.IO.Directory.CreateDirectory(artifactsDir);

        // Save the document.
        doc.Save(System.IO.Path.Combine(artifactsDir, "PositionFootnote.docx"));
    }
}
