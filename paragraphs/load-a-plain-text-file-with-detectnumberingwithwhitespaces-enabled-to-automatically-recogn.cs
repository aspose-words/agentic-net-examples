using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample plain‑text containing list items where numbers are followed by a whitespace.
        const string text = "Shopping list:\n" +
                            "1 Milk\n" +
                            "2 Bread\n" +
                            "3 Eggs\n\n" +
                            "Tasks:\n" +
                            "1 Finish report\n" +
                            "2 Call client\n" +
                            "3 Schedule meeting";

        // Enable detection of list items that use whitespace as a delimiter.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the plain‑text into a Document via a memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
        {
            Document doc = new Document(stream, loadOptions);

            // Output the number of lists detected (for demonstration purposes).
            Console.WriteLine($"Detected lists: {doc.Lists.Count}");

            // Save the resulting Word document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
