using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Settings;

class HyphenationCheck
{
    static void Main()
    {
        // Path to an English hyphenation dictionary (OpenOffice .dic format).
        // If the file does not exist, Aspose.Words will use its built‑in dictionary (if available).
        string dictionaryPath = @"C:\Hyphenation\hyph_en_US.dic";
        if (File.Exists(dictionaryPath))
        {
            // Register the dictionary for the "en-US" locale.
            Hyphenation.RegisterDictionary("en-US", dictionaryPath);
        }

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Set a positive hyphenation zone (zero is not allowed and throws an exception).
        doc.HyphenationOptions.HyphenationZone = 5;
        // Hyphenate words written in all caps as well (not needed for this word).
        doc.HyphenationOptions.HyphenateCaps = true;

        // Build a long paragraph that forces line breaks.
        // Repeating the target word many times ensures it will be split across lines.
        builder.Font.Size = 12;
        for (int i = 0; i < 30; i++)
        {
            builder.Write("hyphenation ");
        }

        // Force the layout engine to calculate pages and lines.
        doc.UpdatePageLayout();

        // Scan the laid‑out lines to see if any line ends with a hyphen.
        // A trailing hyphen indicates that the preceding word was hyphenated.
        bool hyphenated = false;
        LayoutEnumerator enumerator = new LayoutEnumerator(doc);
        while (enumerator.MoveNext())
        {
            if (enumerator.Type == LayoutEntityType.Line)
            {
                string lineText = enumerator.Text;
                if (lineText.EndsWith("-") && lineText.Contains("hyphen"))
                {
                    hyphenated = true;
                    break;
                }
            }
        }

        Console.WriteLine($"Will the word 'hyphenation' be hyphenated? {hyphenated}");
    }
}
