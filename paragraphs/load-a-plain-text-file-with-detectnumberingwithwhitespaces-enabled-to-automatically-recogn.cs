using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample plain‑text content that contains list items using a whitespace delimiter.
        const string textContent =
            "1 First list item 1\n" +
            "2 First list item 2\n" +
            "3 First list item 3\n";

        // Convert the text to a UTF‑8 byte array and load it into a memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(textContent)))
        {
            // Configure load options to recognize list items where the number is followed by a whitespace.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                DetectNumberingWithWhitespaces = true
            };

            // Load the plain‑text document with the specified options.
            Document doc = new Document(stream, loadOptions);

            // The loader should have created a List for the detected items.
            Console.WriteLine($"Number of lists detected: {doc.Lists.Count}");

            // Save the resulting Word document to the local file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "DetectedLists.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
