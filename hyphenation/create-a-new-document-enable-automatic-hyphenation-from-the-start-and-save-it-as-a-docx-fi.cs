using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping where hyphenation can occur.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add a paragraph with long words that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication " +
                        "hyperresponsibility misunderstanding incomprehensibilities " +
                        "characteristically uncharacteristically");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document as DOCX.
        const string outputPath = "HyphenatedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The file '{outputPath}' was not created.");
    }
}
