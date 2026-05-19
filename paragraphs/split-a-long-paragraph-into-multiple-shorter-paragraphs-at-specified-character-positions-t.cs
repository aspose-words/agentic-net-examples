using System;
using Aspose.Words;

namespace SplitParagraphExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a single long paragraph (no explicit line break).
            string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore " +
                              "et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi " +
                              "ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit " +
                              "esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, " +
                              "sunt in culpa qui officia deserunt mollit anim id est laborum.";
            builder.Write(longText); // This creates one paragraph.

            // Retrieve the paragraph that contains the long text.
            Paragraph originalParagraph = doc.FirstSection.Body.FirstParagraph;

            // Get the paragraph text without the trailing paragraph break character.
            string paragraphText = originalParagraph.GetText();
            if (paragraphText.EndsWith(ControlChar.ParagraphBreak))
                paragraphText = paragraphText.Substring(0, paragraphText.Length - ControlChar.ParagraphBreak.Length);

            // Define split positions (character indexes) where new paragraphs should start.
            int[] splitPositions = { 100, 200, 300 }; // Example positions; adjust as needed.

            // Build the list of paragraph fragments.
            var fragments = new System.Collections.Generic.List<string>();
            int start = 0;
            foreach (int pos in splitPositions)
            {
                if (pos > paragraphText.Length) break;
                fragments.Add(paragraphText.Substring(start, pos - start).Trim());
                start = pos;
            }
            if (start < paragraphText.Length)
                fragments.Add(paragraphText.Substring(start).Trim());

            // Remove the original long paragraph.
            originalParagraph.Remove();

            // Move the builder to the end of the document (or any desired location) before inserting new paragraphs.
            builder.MoveToDocumentEnd();

            // Insert the new shorter paragraphs.
            foreach (string fragment in fragments)
            {
                // Write the fragment and end the paragraph.
                builder.Writeln(fragment);
            }

            // Save the resulting document.
            doc.Save("SplitParagraph.docx");
        }
    }
}
