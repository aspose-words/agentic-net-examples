using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Plain‑text source that contains four potential lists.
        // The fourth list uses a whitespace delimiter (e.g. "1 Item").
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

        // Configure loading options – enable whitespace detection so the fourth block
        // is treated as a numbered list. Set to false to disable this behaviour.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the plain‑text into a Word document using the options above.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(textDoc)))
        {
            Document doc = new Document(stream, loadOptions);

            // At this point Aspose.Words has created List objects for each detected list.
            // The count will be 4 when DetectNumberingWithWhitespaces is true,
            // otherwise it will be 3.
            Console.WriteLine($"Number of lists detected: {doc.Lists.Count}");

            // Save the resulting document as DOCX.
            doc.Save("Result.docx");
        }
    }
}
