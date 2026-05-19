using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample plain‑text content that contains list‑like lines.
        const string plainText =
            "Shopping List:\n" +
            "1. Apples\n" +
            "2. Bananas\n" +
            "3. Oranges\n" +
            "\n" +
            "Tasks:\n" +
            "• Finish report\n" +
            "• Call client\n" +
            "• Schedule meeting\n";

        // Load the plain‑text into a Document without automatic list detection.
        var loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = false // keep list items as normal paragraphs.
        };
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(plainText));
        Document doc = new Document(stream, loadOptions);

        // Prepare list objects that will be applied to detected items.
        List numberList = doc.Lists.Add(ListTemplate.NumberDefault);
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Regular expressions for simple numbered and bulleted items.
        Regex numberedRegex = new Regex(@"^\s*\d+[\.\)]\s+");
        Regex bulletRegex = new Regex(@"^\s*[\u2022\u2023\u25E6\u2024\u2025\u2026]\s+"); // common bullet symbols

        // Iterate over all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            string paragraphText = para.GetText(); // includes the trailing paragraph mark.

            // Detect numbered list items.
            if (numberedRegex.IsMatch(paragraphText))
            {
                // Remove the plain‑text number prefix.
                string cleanedText = numberedRegex.Replace(paragraphText, string.Empty);
                // Replace the paragraph's runs with a single run containing the cleaned text.
                para.Runs.Clear();
                para.AppendChild(new Run(doc, cleanedText));
                // Apply the numbered list formatting.
                para.ListFormat.List = numberList;
                para.ListFormat.ListLevelNumber = 0;
            }
            // Detect bulleted list items.
            else if (bulletRegex.IsMatch(paragraphText))
            {
                string cleanedText = bulletRegex.Replace(paragraphText, string.Empty);
                para.Runs.Clear();
                para.AppendChild(new Run(doc, cleanedText));
                para.ListFormat.List = bulletList;
                para.ListFormat.ListLevelNumber = 0;
            }
        }

        // Save the resulting document.
        doc.Save("ConvertedLists.docx");
    }
}
