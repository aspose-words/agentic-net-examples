using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample plain‑text containing list items.
        // The fourth list uses a whitespace delimiter, which is recognized only when
        // DetectNumberingWithWhitespaces is set to true.
        const string text = 
            "Full stop delimiters:\n" +
            "1. First list item 1\n" +
            "2. First list item 2\n" +
            "3. First list item 3\n\n" +
            "Whitespace delimiters:\n" +
            "1 Fourth list item 1\n" +
            "2 Fourth list item 2\n" +
            "3 Fourth list item 3";

        // Configure load options to detect list items that use whitespaces as delimiters.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the plain‑text into a Document using a memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
        {
            Document doc = new Document(stream, loadOptions);

            // Optional: display the number of detected lists.
            Console.WriteLine($"Detected lists: {doc.Lists.Count}");

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DetectedLists.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
