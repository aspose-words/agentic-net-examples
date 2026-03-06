using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Plain‑text source containing several list styles.
        string text = "Full stop delimiters:\n" +
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

        // Configure loading options – enable detection of list items that use whitespace as a delimiter.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.DetectNumberingWithWhitespaces = true; // Set to false to ignore whitespace‑delimited lists.

        // Load the plain‑text into a Document using the configured options.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
        {
            Document doc = new Document(stream, loadOptions);

            // Display how many lists were recognized.
            Console.WriteLine($"Lists detected: {doc.Lists.Count}");

            // Save the resulting DOCX file.
            doc.Save("Result.docx");
        }
    }
}
