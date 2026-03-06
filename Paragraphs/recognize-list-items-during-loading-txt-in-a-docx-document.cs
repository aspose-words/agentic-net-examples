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
        // Plain‑text source containing a list that uses a whitespace delimiter.
        const string text = "Full stop delimiters:\n" +
                            "1. First list item 1\n" +
                            "2. First list item 2\n" +
                            "3. First list item 3\n\n" +
                            "Whitespace delimiters:\n" +
                            "1 Fourth list item 1\n" +
                            "2 Fourth list item 2\n" +
                            "3 Fourth list item 3";

        // Load with whitespace detection enabled.
        Document docWithWs = LoadTxtDocument(text, detectWhitespace: true);
        Console.WriteLine($"Lists detected (whitespace enabled): {docWithWs.Lists.Count}");
        bool hasWsList = docWithWs.FirstSection.Body.Paragraphs
            .Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem);
        Console.WriteLine($"Fourth list recognized as list item: {hasWsList}");

        // Load with whitespace detection disabled.
        Document docWithoutWs = LoadTxtDocument(text, detectWhitespace: false);
        Console.WriteLine($"Lists detected (whitespace disabled): {docWithoutWs.Lists.Count}");
        bool hasWsListDisabled = docWithoutWs.FirstSection.Body.Paragraphs
            .Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem);
        Console.WriteLine($"Fourth list recognized as list item: {hasWsListDisabled}");

        // Save the results for visual verification.
        docWithWs.Save("WithWhitespace.docx");
        docWithoutWs.Save("WithoutWhitespace.docx");
    }

    static Document LoadTxtDocument(string txt, bool detectWhitespace)
    {
        // Convert the plain text to a UTF‑8 memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(txt)))
        {
            // Configure load options.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                DetectNumberingWithWhitespaces = detectWhitespace
            };

            // Load the document from the stream using the specified options.
            return new Document(stream, loadOptions);
        }
    }
}
