using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPaginationDemo
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a multi‑section document with narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page setup.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add first section content.
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Insert a section break and add second section content.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Ensure layout is calculated.
        doc.UpdatePageLayout();
        int pageCountWithoutHyphenation = doc.PageCount;

        // Save the document without hyphenation.
        const string withoutHyphenationFile = "ReportWithoutHyphenation.pdf";
        doc.Save(withoutHyphenationFile, SaveFormat.Pdf);
        if (!File.Exists(withoutHyphenationFile))
            throw new InvalidOperationException("Failed to create the PDF without hyphenation.");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Re‑calculate layout after enabling hyphenation.
        doc.UpdatePageLayout();
        int pageCountWithHyphenation = doc.PageCount;

        // Save the document with hyphenation.
        const string withHyphenationFile = "ReportWithHyphenation.pdf";
        doc.Save(withHyphenationFile, SaveFormat.Pdf);
        if (!File.Exists(withHyphenationFile))
            throw new InvalidOperationException("Failed to create the PDF with hyphenation.");

        // Output the pagination comparison.
        Console.WriteLine($"Pages without hyphenation: {pageCountWithoutHyphenation}");
        Console.WriteLine($"Pages with hyphenation   : {pageCountWithHyphenation}");
    }
}
