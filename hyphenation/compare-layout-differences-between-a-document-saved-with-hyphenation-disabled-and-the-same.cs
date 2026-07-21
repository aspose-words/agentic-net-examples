using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings; // For HyphenationOptions

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string outputDir = Path.Combine(Path.GetTempPath(), "HyphenationDemo");
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (en-US).
        // The dictionary must contain hyphenation points for the words we will use.
        string dictPath = Path.Combine(outputDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a document with a narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing long words that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // -----------------------------------------------------------------
        // Save version with hyphenation disabled (default).
        // -----------------------------------------------------------------
        string disabledPath = Path.Combine(outputDir, "HyphenationDisabled.pdf");
        doc.Save(disabledPath, SaveFormat.Pdf);
        if (!File.Exists(disabledPath))
            throw new InvalidOperationException("Failed to create disabled hyphenation file.");

        // -----------------------------------------------------------------
        // Enable automatic hyphenation and save the second version.
        // -----------------------------------------------------------------
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        string enabledPath = Path.Combine(outputDir, "HyphenationEnabled.pdf");
        doc.Save(enabledPath, SaveFormat.Pdf);
        if (!File.Exists(enabledPath))
            throw new InvalidOperationException("Failed to create enabled hyphenation file.");

        // -----------------------------------------------------------------
        // Load both PDFs to compare layout (page count).
        // -----------------------------------------------------------------
        Document disabledDoc = new Document(disabledPath);
        Document enabledDoc = new Document(enabledPath);

        // Ensure layout is up‑to‑date.
        disabledDoc.UpdatePageLayout();
        enabledDoc.UpdatePageLayout();

        int disabledPages = disabledDoc.PageCount;
        int enabledPages = enabledDoc.PageCount;

        Console.WriteLine($"Hyphenation disabled pages: {disabledPages}");
        Console.WriteLine($"Hyphenation enabled pages: {enabledPages}");

        if (disabledPages != enabledPages)
        {
            Console.WriteLine("Layout differs: page count changed due to hyphenation.");
        }
        else
        {
            Console.WriteLine("Layout does not differ in page count.");
        }

        // Clean up temporary files (optional).
        // Comment out the following lines if you want to inspect the PDFs.
        try
        {
            File.Delete(dictPath);
            File.Delete(disabledPath);
            File.Delete(enabledPath);
            Directory.Delete(outputDir);
        }
        catch
        {
            // Ignored – cleanup is best‑effort.
        }
    }
}
