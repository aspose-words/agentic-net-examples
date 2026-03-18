using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationPageCountDemo
{
    static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a long report to force many pages.
        builder.Font.Size = 12;
        for (int i = 0; i < 200; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        }

        // -------------------- Without hyphenation --------------------
        doc.HyphenationOptions.AutoHyphenation = false; // ensure hyphenation is off
        doc.UpdatePageLayout();                         // rebuild layout
        int pagesWithoutHyphenation = doc.PageCount;    // total pages
        doc.Save("ReportWithoutHyphenation.docx");      // save the document

        // -------------------- With hyphenation --------------------
        // Clone the original content so the two files are independent.
        Document hyphenatedDoc = (Document)doc.Clone();
        hyphenatedDoc.HyphenationOptions.AutoHyphenation = true; // enable hyphenation
        hyphenatedDoc.HyphenationOptions.HyphenationZone = 720;   // 0.5 inch from right margin
        hyphenatedDoc.UpdatePageLayout();                         // rebuild layout
        int pagesWithHyphenation = hyphenatedDoc.PageCount;       // total pages
        hyphenatedDoc.Save("ReportWithHyphenation.docx");        // save the hyphenated version

        // Output the page counts for comparison.
        Console.WriteLine($"Pages without hyphenation: {pagesWithoutHyphenation}");
        Console.WriteLine($"Pages with hyphenation: {pagesWithHyphenation}");
    }
}
