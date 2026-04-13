using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPaginationExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a multi‑section document with long paragraphs that will wrap.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to increase the chance of hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the language to English (US) – the default hyphenation dictionary is available for this locale.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Add first section content.
        builder.Writeln("Section 1");
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ";
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln(longText);
        }

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln(longText);
        }

        // Clone the document before enabling hyphenation – this will be the “without hyphenation” version.
        Document docWithoutHyphen = (Document)doc.Clone(true);
        docWithoutHyphen.UpdatePageLayout();
        int pagesWithoutHyphen = docWithoutHyphen.PageCount;

        // Save the version without hyphenation.
        string withoutHyphenPath = Path.Combine(outputDir, "ReportWithoutHyphenation.pdf");
        docWithoutHyphen.Save(withoutHyphenPath, SaveFormat.Pdf);

        // Enable automatic hyphenation on the original document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // HyphenationZone must be a non‑negative value; using the default (360) which equals 0.25 inch.
        doc.HyphenationOptions.HyphenationZone = 360;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Re‑layout the document after changing hyphenation settings.
        doc.UpdatePageLayout();
        int pagesWithHyphen = doc.PageCount;

        // Save the hyphenated version.
        string withHyphenPath = Path.Combine(outputDir, "ReportWithHyphenation.pdf");
        doc.Save(withHyphenPath, SaveFormat.Pdf);

        // Output the pagination comparison.
        Console.WriteLine($"Pages without hyphenation: {pagesWithoutHyphen}");
        Console.WriteLine($"Pages with hyphenation   : {pagesWithHyphen}");
        Console.WriteLine($"Difference               : {pagesWithoutHyphen - pagesWithHyphen}");
        Console.WriteLine($"Outputs saved to: {outputDir}");
    }
}
