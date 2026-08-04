using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample plain‑text containing list items separated by whitespace.
        const string text = 
            "Shopping list:\n" +
            "1 Milk\n" +
            "2 Bread\n" +
            "3 Eggs\n\n" +
            "Tasks:\n" +
            "1 Finish report\n" +
            "2 Call client\n" +
            "3 Schedule meeting";

        // Prepare load options with whitespace‑based list detection enabled.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the text into a Word document using a memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
        {
            Document doc = new Document(stream, loadOptions);

            // Output the number of detected lists (should be 2 in this example).
            Console.WriteLine($"Detected lists: {doc.Lists.Count}");

            // Verify that a paragraph from the second list is recognized as a list item.
            bool isListItem = doc.FirstSection.Body.Paragraphs
                .Any(p => p.GetText().Contains("Finish report") && ((Paragraph)p).IsListItem);
            Console.WriteLine($"\"Finish report\" is a list item: {isListItem}");

            // Save the resulting document.
            doc.Save("Output.docx");
        }
    }
}
