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

        // Set a narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a long paragraph that will require hyphenation.
        builder.Font.Size = 12;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        builder.Writeln(
            "Aspose.Words provides powerful hyphenation features that allow automatic breaking of long words " +
            "across lines, improving the visual appearance of justified text in narrow columns.");

        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "Aspose.Words=As-pose.Words\n" +
            "features=fea-tures\n" +
            "automatic=au-to-matic\n" +
            "justified=jus-ti-fied\n");

        // Register the dictionary and enable automatic hyphenation.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // increase zone to reduce gaps

        // Use compression justification to tighten spacing after hyphenation.
        doc.JustificationMode = JustificationMode.Compress;

        // Save the document to PDF.
        const string outputPath = "HyphenatedCompressed.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
