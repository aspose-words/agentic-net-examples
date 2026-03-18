using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Plain‑text content that contains numbered list items.
        const string plainText =
            "Full stop delimiters:\n" +
            "1. First item\n" +
            "2. Second item\n" +
            "3. Third item\n\n" +
            "Whitespace delimiters:\n" +
            "1 Fourth item\n" +
            "2 Fifth item\n" +
            "3 Sixth item";

        // Prepare a memory stream with the UTF‑8 encoded text.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(plainText)))
        {
            // Configure load options to recognise list numbers that are separated by whitespaces.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                // Enable detection of list items that use a whitespace as the delimiter (e.g., "1 Item").
                DetectNumberingWithWhitespaces = true,
                // Keep the default automatic numbering detection (true) – optional as it is true by default.
                AutoNumberingDetection = true
            };

            // Load the plain‑text document. The constructor will apply the list detection algorithm.
            Document doc = new Document(stream, loadOptions);

            // Ensure that list labels are up‑to‑date (useful if further processing modifies the document).
            doc.UpdateListLabels();

            // Example: output the number of detected lists.
            Console.WriteLine($"Detected lists: {doc.Lists.Count}");

            // Save the resulting Word document.
            doc.Save("DetectedLists.docx");
        }
    }
}
