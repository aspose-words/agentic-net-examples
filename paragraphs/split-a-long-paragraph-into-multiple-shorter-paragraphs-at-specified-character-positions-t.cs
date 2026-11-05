using System;
using System.Collections.Generic;
using Aspose.Words;

namespace SplitParagraphExample
{
    public class Program
    {
        public static void Main()
        {
            // Sample long paragraph text.
            string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                              "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                              "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                              "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in " +
                              "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
                              "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in " +
                              "culpa qui officia deserunt mollit anim id est laborum.";

            // Character positions where the paragraph should be split.
            // Positions are zero‑based indices in the string.
            int[] splitPositions = { 80, 160, 240 };

            // Ensure split positions are sorted and within the text length.
            List<int> positions = new List<int>();
            foreach (int pos in splitPositions)
            {
                if (pos > 0 && pos < longText.Length)
                    positions.Add(pos);
            }
            positions.Sort();

            // Create substrings based on split positions.
            List<string> parts = new List<string>();
            int start = 0;
            foreach (int pos in positions)
            {
                int length = pos - start;
                if (length > 0)
                {
                    parts.Add(longText.Substring(start, length).Trim());
                    start = pos;
                }
            }
            // Add the remaining part.
            if (start < longText.Length)
                parts.Add(longText.Substring(start).Trim());

            // Create a new document and insert the split paragraphs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            foreach (string part in parts)
            {
                builder.Writeln(part);
            }

            // Save the resulting document.
            doc.Save("SplitParagraph.docx");
        }
    }
}
