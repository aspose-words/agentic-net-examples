using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add sample text that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Insert a second section with the same content.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        // Ensure the second section has the same page setup.
        doc.Sections[1].PageSetup.PageWidth = doc.FirstSection.PageSetup.PageWidth;
        doc.Sections[1].PageSetup.LeftMargin = doc.FirstSection.PageSetup.LeftMargin;
        doc.Sections[1].PageSetup.RightMargin = doc.FirstSection.PageSetup.RightMargin;

        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Layout the document and get the page count before hyphenation.
        doc.UpdatePageLayout();
        int pagesBefore = doc.PageCount;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Register the dictionary for the document's language.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Re‑layout the document and get the new page count.
        doc.UpdatePageLayout();
        int pagesAfter = doc.PageCount;

        // Save the hyphenated document as PDF.
        const string outputPath = "Hyphenated.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output file was not created.");

        // Output the pagination comparison.
        Console.WriteLine($"Pages before hyphenation: {pagesBefore}");
        Console.WriteLine($"Pages after hyphenation: {pagesAfter}");
    }
}
