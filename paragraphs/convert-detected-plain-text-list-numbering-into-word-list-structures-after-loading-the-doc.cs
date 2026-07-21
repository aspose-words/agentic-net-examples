using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Lists;   // Needed for List and ListTemplate types

namespace ListConversionExample
{
    public class Program
    {
        public static void Main()
        {
            // Plain‑text content that contains numbered list items.
            const string plainText =
                "1. First item\r\n" +
                "2. Second item\r\n" +
                "3. Third item\r\n" +
                "\r\n" +
                "This is a normal paragraph without list numbering.";

            // Load the plain‑text into a Document, enabling list detection.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                // Detect numbering even when the delimiter is a whitespace.
                DetectNumberingWithWhitespaces = true
            };

            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(plainText)))
            {
                // Create the document from the memory stream using the load options.
                Document doc = new Document(stream, loadOptions);

                // Create a numbered list that will be applied to detected items.
                List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

                // Regular expression to identify lines that start with a number followed by '.' or ')'.
                Regex listItemRegex = new Regex(@"^\d+[\.\)]\s+", RegexOptions.Compiled);

                // Traverse all paragraphs in the document.
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph paragraph in paragraphs)
                {
                    // Trim the paragraph text (removes the trailing paragraph mark).
                    string paragraphText = paragraph.GetText().Trim();

                    // If the paragraph is not already a list item and matches the pattern, convert it.
                    if (!paragraph.ListFormat.IsListItem && listItemRegex.IsMatch(paragraphText))
                    {
                        // Apply the list formatting.
                        paragraph.ListFormat.List = numberedList;
                        paragraph.ListFormat.ListLevelNumber = 0;

                        // Remove the leading number from the paragraph text.
                        string cleanedText = listItemRegex.Replace(paragraphText, string.Empty);

                        // Replace the paragraph's runs with a single run containing the cleaned text.
                        paragraph.Runs.Clear();
                        Run run = new Run(doc, cleanedText);
                        paragraph.AppendChild(run);
                    }
                }

                // Save the resulting document.
                const string outputPath = "ConvertedList.docx";
                doc.Save(outputPath);
                Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
            }
        }
    }
}
