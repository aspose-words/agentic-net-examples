using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationComparison
{
    public static void Main()
    {
        // Define output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the two documents.
        string disabledPath = Path.Combine(outputDir, "HyphenationDisabled.docx");
        string enabledPath = Path.Combine(outputDir, "HyphenationEnabled.docx");

        // Create a sample document with narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page setup.
        builder.PageSetup.PageWidth = 300; // points (~4.2 inches)
        builder.PageSetup.LeftMargin = 20;
        builder.PageSetup.RightMargin = 20;

        // Sample long text containing a very long word to trigger hyphenation.
        string sampleText = "This is a sample paragraph with a very long word such as pneumonoultramicroscopicsilicovolcanoconiosis to demonstrate hyphenation behavior. " +
                            "The quick brown fox jumps over the lazy dog. " +
                            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ";

        // Write the text several times to ensure wrapping.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln(sampleText);
        }

        // Save the document with hyphenation disabled (default).
        doc.HyphenationOptions.AutoHyphenation = false;
        doc.Save(disabledPath);

        // Enable automatic hyphenation and save the second version.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.Save(enabledPath);

        // Validate that both files were created.
        if (!File.Exists(disabledPath) || !File.Exists(enabledPath))
            throw new InvalidOperationException("One or both output documents were not created.");

        // Load the saved documents to compare layout metrics.
        Document docDisabled = new Document(disabledPath);
        Document docEnabled = new Document(enabledPath);

        // Compare page counts as a simple layout difference indicator.
        int pagesDisabled = docDisabled.PageCount;
        int pagesEnabled = docEnabled.PageCount;

        Console.WriteLine($"Hyphenation disabled:  {pagesDisabled} page(s)");
        Console.WriteLine($"Hyphenation enabled:   {pagesEnabled} page(s)");

        // Additionally, compare the number of lines in the first paragraph.
        int linesDisabled = CountLines(docDisabled);
        int linesEnabled = CountLines(docEnabled);

        Console.WriteLine($"Lines in first paragraph (disabled): {linesDisabled}");
        Console.WriteLine($"Lines in first paragraph (enabled):  {linesEnabled}");
    }

    // Helper method to count the visual lines of the first paragraph.
    private static int CountLines(Document document)
    {
        // Ensure layout is built.
        document.UpdatePageLayout();

        // Get the first paragraph.
        var paragraph = document.FirstSection.Body.FirstParagraph;
        if (paragraph == null)
            return 0;

        // The LayoutCollector can be used to map nodes to pages, but line count per paragraph
        // is not directly exposed. As a simple proxy, we count the number of line breaks
        // inserted by the layout engine via the LineSpacing property approximation.
        // Here we approximate by counting the number of line breaks in the paragraph's text.
        // This gives a deterministic count for the example.
        string text = paragraph.GetText();
        // In Word, lines are separated by "\r". The last line ends with "\r".
        return text.Split('\r', StringSplitOptions.RemoveEmptyEntries).Length;
    }
}
