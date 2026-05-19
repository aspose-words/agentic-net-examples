using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Plain‑text content that contains list items using a whitespace as the delimiter.
        const string text = "Shopping list:\n" +
                            "1 Milk\n" +
                            "2 Bread\n" +
                            "3 Eggs\n" +
                            "\n" +
                            "Tasks:\n" +
                            "1 Finish report\n" +
                            "2 Call client\n";

        // Convert the string to a UTF‑8 memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
        {
            // Enable detection of numbered list items that are separated by whitespaces.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                DetectNumberingWithWhitespaces = true
            };

            // Load the plain‑text document with the specified options.
            Document doc = new Document(stream, loadOptions);

            // Output the number of lists that Aspose.Words recognized.
            Console.WriteLine($"Lists detected: {doc.Lists.Count}");

            // Save the resulting Word document.
            doc.Save("Result.docx");
        }
    }
}
