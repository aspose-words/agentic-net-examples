using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Sample text that contains long words to trigger hyphenation.
        string longText = "Antidisestablishmentarianism is often cited as one of the longest words in the English language, " +
                          "and it provides a good example for testing automatic hyphenation in a narrow column. " +
                          "Supercalifragilisticexpialidocious is another famously long word that can be hyphenated.";

        // Create two documents – one with hyphenation enabled and one without.
        Document docWithHyphenation = CreateDocument(longText, enableHyphenation: true);
        Document docWithoutHyphenation = CreateDocument(longText, enableHyphenation: false);

        // Measure PDF generation time with hyphenation enabled.
        string pdfWithHyphenationPath = Path.Combine(outputDir, "HyphenationEnabled.pdf");
        var sw = Stopwatch.StartNew();
        docWithHyphenation.Save(pdfWithHyphenationPath, SaveFormat.Pdf);
        sw.Stop();
        long timeWithHyphenation = sw.ElapsedMilliseconds;

        // Measure PDF generation time with hyphenation disabled.
        string pdfWithoutHyphenationPath = Path.Combine(outputDir, "HyphenationDisabled.pdf");
        sw.Restart();
        docWithoutHyphenation.Save(pdfWithoutHyphenationPath, SaveFormat.Pdf);
        sw.Stop();
        long timeWithoutHyphenation = sw.ElapsedMilliseconds;

        // Output the measured times.
        Console.WriteLine($"PDF generation time with hyphenation enabled : {timeWithHyphenation} ms");
        Console.WriteLine($"PDF generation time with hyphenation disabled: {timeWithoutHyphenation} ms");
        Console.WriteLine($"Generated files are located in: {outputDir}");
    }

    // Creates a document containing the supplied text and configures hyphenation.
    private static Document CreateDocument(string text, bool enableHyphenation)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Set a narrow page width to increase the chance of hyphenation.
        // Page width is measured in points (1 inch = 72 points).
        builder.PageSetup.PageWidth = 300; // ~4.17 inches.
        builder.PageSetup.LeftMargin = 20; // ~0.28 inch.
        builder.PageSetup.RightMargin = 20; // ~0.28 inch.

        // Use a readable font size.
        builder.Font.Size = 12;

        // Set the language/locale for hyphenation patterns.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write the sample text.
        builder.Writeln(text);

        // Configure hyphenation options.
        doc.HyphenationOptions.AutoHyphenation = enableHyphenation;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (in 1/20 point units).
        doc.HyphenationOptions.HyphenateCaps = true;

        return doc;
    }
}
