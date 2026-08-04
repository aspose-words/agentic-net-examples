using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPdfExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a long paragraph that will require hyphenation when the page width is narrow.
        builder.Font.Size = 24;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "demonstration of automatic hyphenation in a narrow column.");

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "demonstration=dem-on-stra-tion\n" +
            "automatic=au-to-ma-tic\n" +
            "hyphenation=hy-phen-a-tion\n" +
            "narrow=nar-row\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: limit consecutive hyphenated lines.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // Ensure hyphenation marks are not rendered as visible characters in the PDF.
        // In Aspose.Words, automatic hyphenation inserts soft hyphens that are not shown in PDF output.
        // No additional configuration is required beyond enabling AutoHyphenation.

        // Save the document as PDF.
        const string pdfFileName = "HyphenatedOutput.pdf";
        doc.Save(pdfFileName);

        // Validate that the PDF was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("The PDF file was not created.");

        Console.WriteLine($"PDF generated successfully: {Path.GetFullPath(pdfFileName)}");
    }
}
