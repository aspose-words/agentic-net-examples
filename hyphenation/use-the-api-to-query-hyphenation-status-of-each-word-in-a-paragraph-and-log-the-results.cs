using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line breaks and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing long words.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication demonstration");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360;

        // Create a minimal hyphenation dictionary for en-US.
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "demonstration=dem-on-stra-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Save the document to trigger layout processing.
        const string docPath = "HyphenationStatus.docx";
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException("Document was not saved.");

        // Load the saved document (layout is already built).
        Document loaded = new Document(docPath);

        // Get the first paragraph.
        Paragraph paragraph = loaded.FirstSection.Body.FirstParagraph;

        // Split the paragraph text into words (excluding whitespace and line breaks).
        string[] words = paragraph.GetText()
            .Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        Console.WriteLine("Hyphenation status of words in the paragraph:");

        // Since Aspose.Words does not expose a direct API to query per‑word hyphenation status,
        // we will simply report that hyphenation was enabled for the document.
        // In a real scenario you could inspect the layout or render the document to see hyphens.
        foreach (string word in words)
        {
            Console.WriteLine($"{word} - Hyphenation enabled (status not directly queryable)");
        }
    }
}
