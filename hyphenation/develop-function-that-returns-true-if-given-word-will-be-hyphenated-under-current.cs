using System;
using Aspose.Words;

namespace HyphenationDemo
{
    public static class HyphenationHelper
    {
        /// <summary>
        /// Returns true if the specified word can be hyphenated under the current hyphenation settings.
        /// The method creates a temporary document, enables automatic hyphenation,
        /// forces a narrow line width to trigger hyphenation and checks for the presence
        /// of a soft‑hyphen character (U+00AD) after layout.
        /// </summary>
        public static bool IsWordHyphenated(string word)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write the word into the document.
            builder.Font.Size = 12;
            builder.Writeln(word);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Reduce the page width so that the word does not fit on a single line,
            // forcing the layout engine to consider hyphenation.
            doc.FirstSection.PageSetup.PageWidth = 50; // 50 points (~0.7 inch)

            // Force layout to be performed.
            doc.UpdatePageLayout();

            // Retrieve the rendered text of the first paragraph.
            string renderedText = doc.FirstSection.Body.FirstParagraph.GetText();

            // Hyphenation inserts a soft hyphen character (U+00AD) into the text.
            return renderedText.Contains('\u00AD');
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – you can replace the word with any test case.
            string testWord = args.Length > 0 ? args[0] : "hyphenation";
            bool hyphenated = HyphenationHelper.IsWordHyphenated(testWord);
            Console.WriteLine($"Word '{testWord}' hyphenated: {hyphenated}");
        }
    }
}
