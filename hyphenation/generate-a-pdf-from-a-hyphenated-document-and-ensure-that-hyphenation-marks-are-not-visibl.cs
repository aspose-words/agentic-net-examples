using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Path for the temporary hyphenation dictionary.
        const string dictPath = "hyph_en_US.dic";

        // Create a minimal hyphenation dictionary for English (US).
        // The dictionary format is the OpenOffice hyphenation format.
        // It contains a header line followed by word patterns.
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate words of the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping and thus hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph that contains long words present in the dictionary.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Suppress the visual hyphenation marks in the output.
        // This keeps the layout hyphenated but does not render hyphens.
        builder.ParagraphFormat.SuppressAutoHyphens = true;

        // Save the document as PDF.
        const string pdfPath = "hyphenated.pdf";
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictPath))
            File.Delete(dictPath);
    }
}
