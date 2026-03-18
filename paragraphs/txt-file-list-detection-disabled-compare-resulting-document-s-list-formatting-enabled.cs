using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Sample plain‑text containing several list styles.
        const string textDoc =
            "Full stop delimiters:\n" +
            "1. First list item 1\n" +
            "2. First list item 2\n" +
            "3. First list item 3\n\n" +
            "Right bracket delimiters:\n" +
            "1) Second list item 1\n" +
            "2) Second list item 2\n" +
            "3) Second list item 3\n\n" +
            "Bullet delimiters:\n" +
            "• Third list item 1\n" +
            "• Third list item 2\n" +
            "• Third list item 3\n\n" +
            "Whitespace delimiters:\n" +
            "1 Fourth list item 1\n" +
            "2 Fourth list item 2\n" +
            "3 Fourth list item 3";

        // Helper that loads the text into a Document using the specified option.
        Document LoadDocument(bool detectNumberingWithWhitespaces)
        {
            // Configure TxtLoadOptions.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                DetectNumberingWithWhitespaces = detectNumberingWithWhitespaces
            };

            // Load from a memory stream using the Document constructor that accepts a stream and LoadOptions.
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(textDoc)))
            {
                return new Document(stream, loadOptions);
            }
        }

        // Load twice: once with detection enabled (true) and once disabled (false).
        Document docEnabled = LoadDocument(true);
        Document docDisabled = LoadDocument(false);

        // Output the number of List objects detected in each mode.
        Console.WriteLine($"Lists count with detection enabled : {docEnabled.Lists.Count}");
        Console.WriteLine($"Lists count with detection disabled: {docDisabled.Lists.Count}");

        // Determine whether the "Fourth list" paragraphs were recognized as list items.
        bool fourthListEnabled = docEnabled.FirstSection.Body.Paragraphs
            .Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem);

        bool fourthListDisabled = docDisabled.FirstSection.Body.Paragraphs
            .Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem);

        Console.WriteLine($"Fourth list recognized with detection enabled : {fourthListEnabled}");
        Console.WriteLine($"Fourth list recognized with detection disabled: {fourthListDisabled}");
    }
}
