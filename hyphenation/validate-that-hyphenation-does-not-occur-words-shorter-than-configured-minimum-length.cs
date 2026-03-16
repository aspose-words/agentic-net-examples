using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationValidation
{
    static void Main()
    {
        // Minimum word length that should be allowed to hyphenate.
        const int minHyphenationLength = 5;

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: set hyphenation zone to avoid hyphenation near the right margin.
        doc.HyphenationOptions.HyphenationZone = 720;
        // Optional: allow hyphenation of capitalized words.
        doc.HyphenationOptions.HyphenateCaps = true;

        // Sample text containing short and long words.
        string text = "A quick brown fox jumps over the lazy dog. " +
                      "Supercalifragilisticexpialidocious is a long word.";

        // Write the text into the document.
        builder.Writeln(text);

        // Iterate over all paragraphs in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Split the paragraph text into words.
            string[] words = paragraph.GetText().Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Determine if the paragraph contains any word shorter than the minimum length.
            bool containsShortWord = false;
            foreach (string word in words)
            {
                // Trim punctuation.
                string trimmed = word.Trim(new char[] { '.', ',', ';', ':', '!', '?' });
                if (trimmed.Length > 0 && trimmed.Length < minHyphenationLength)
                {
                    containsShortWord = true;
                    break;
                }
            }

            // If the paragraph has short words, suppress automatic hyphenation for it.
            if (containsShortWord)
                paragraph.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Save the document.
        doc.Save("HyphenationValidation.docx");
    }
}
