using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Sample long text that will cause many line wraps.
        string longText = string.Concat(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, ",
            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ",
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ",
            "ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit ",
            "in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ",
            "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia ",
            "deserunt mollit anim id est laborum. ");

        // Repeat the paragraph many times to ensure the document spans several pages.
        int repeatCount = 50;
        string reportContent = string.Concat(Enumerable.Repeat(longText, repeatCount));

        // -----------------------------------------------------------------
        // 1. Document without hyphenation (default AutoHyphenation = false).
        // -----------------------------------------------------------------
        Document docNoHyphen = new Document();
        ConfigurePageSetup(docNoHyphen);
        DocumentBuilder builderNoHyphen = new DocumentBuilder(docNoHyphen);
        builderNoHyphen.Writeln(reportContent);
        string noHyphenPath = Path.Combine(outputDir, "Report_NoHyphenation.docx");
        docNoHyphen.Save(noHyphenPath);
        int pagesNoHyphen = docNoHyphen.PageCount;

        // -----------------------------------------------------------------
        // 2. Document with automatic hyphenation enabled.
        // -----------------------------------------------------------------
        Document docHyphen = new Document();
        ConfigurePageSetup(docHyphen);
        // Enable automatic hyphenation.
        docHyphen.HyphenationOptions.AutoHyphenation = true;
        // Optional: fine‑tune hyphenation behaviour.
        docHyphen.HyphenationOptions.ConsecutiveHyphenLimit = 0; // 0 = unlimited consecutive hyphens
        docHyphen.HyphenationOptions.HyphenationZone = 360;      // Default value (0.25 inch) – must be > 0
        docHyphen.HyphenationOptions.HyphenateCaps = true;

        DocumentBuilder builderHyphen = new DocumentBuilder(docHyphen);
        builderHyphen.Writeln(reportContent);
        string hyphenPath = Path.Combine(outputDir, "Report_WithHyphenation.docx");
        docHyphen.Save(hyphenPath);
        int pagesHyphen = docHyphen.PageCount;

        // -----------------------------------------------------------------
        // Validation and output.
        // -----------------------------------------------------------------
        Console.WriteLine($"Pages without hyphenation: {pagesNoHyphen}");
        Console.WriteLine($"Pages with hyphenation   : {pagesHyphen}");

        if (!File.Exists(noHyphenPath))
            throw new FileNotFoundException("The document without hyphenation was not created.", noHyphenPath);
        if (!File.Exists(hyphenPath))
            throw new FileNotFoundException("The document with hyphenation was not created.", hyphenPath);

        // Hyphenation should not increase the page count; it usually reduces it.
        if (pagesHyphen > pagesNoHyphen)
            throw new InvalidOperationException("Hyphenation increased the page count, which is unexpected for this test.");

        // Example completed successfully.
    }

    // Configures a narrow page width and small margins to force line wrapping.
    private static void ConfigurePageSetup(Document doc)
    {
        // 4.5 inches width (approx 324 points) minus margins.
        const double pageWidthPoints = 324.0;
        const double marginPoints = 20.0;

        doc.FirstSection.PageSetup.PageWidth = pageWidthPoints;
        doc.FirstSection.PageSetup.LeftMargin = marginPoints;
        doc.FirstSection.PageSetup.RightMargin = marginPoints;
        doc.FirstSection.PageSetup.TopMargin = marginPoints;
        doc.FirstSection.PageSetup.BottomMargin = marginPoints;
    }
}
