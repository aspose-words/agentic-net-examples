using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Plain‑text source that contains list items using whitespace as the delimiter.
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

        // Convert the string to a UTF‑8 memory stream – this is the source for the Document constructor.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(textDoc)))
        {
            // Configure loading options so that whitespace delimiters are treated as list markers.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                DetectNumberingWithWhitespaces = true
            };

            // Load the plain‑text into a Word document using the specified options.
            Document doc = new Document(stream, loadOptions);

            // At this point Aspose.Words has created List objects for each detected list.
            // For demonstration we can output the number of lists detected.
            Console.WriteLine($"Lists detected: {doc.Lists.Count}");

            // Save the resulting document as DOCX.
            doc.Save("RecognizedLists.docx");
        }
    }
}
