using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationStatusLogger
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample text.
        builder.Writeln("Hyphenation allows words to be split across lines, improving text justification.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // optional, just to illustrate usage.

        // Get the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Determine if this paragraph is exempt from hyphenation.
        bool paragraphSuppressesHyphens = paragraph.ParagraphFormat.SuppressAutoHyphens;

        // Split the paragraph text into words.
        string paragraphText = paragraph.GetText(); // Includes trailing paragraph break.
        string[] words = paragraphText.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        // Log hyphenation status for each word.
        Console.WriteLine("Hyphenation status per word:");
        foreach (string word in words)
        {
            // Hyphenation is possible if document auto‑hyphenation is on and the paragraph does not suppress it.
            bool hyphenationEnabled = doc.HyphenationOptions.AutoHyphenation && !paragraphSuppressesHyphens;

            // For demonstration we also consider words in all caps (controlled by HyphenateCaps).
            bool canHyphenateCaps = doc.HyphenationOptions.HyphenateCaps || !IsAllCaps(word);

            bool canHyphenate = hyphenationEnabled && canHyphenateCaps;

            Console.WriteLine($"Word: \"{word}\"  Hyphenation allowed: {canHyphenate}");
        }

        // Save the document (required by lifecycle rule).
        doc.Save("HyphenationStatus.docx");
    }

    // Helper to detect all‑capital words.
    private static bool IsAllCaps(string word)
    {
        foreach (char ch in word)
        {
            if (char.IsLetter(ch) && !char.IsUpper(ch))
                return false;
        }
        return true;
    }
}
