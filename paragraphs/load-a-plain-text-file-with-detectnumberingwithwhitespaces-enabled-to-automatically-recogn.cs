using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Plain‑text content that contains list items separated by whitespace.
        const string text = "Shopping list:\n" +
                            "1 Milk\n" +
                            "2 Bread\n" +
                            "3 Eggs\n\n" +
                            "Tasks:\n" +
                            "1 Finish report\n" +
                            "2 Call client\n";

        // Configure load options to treat whitespace as a list delimiter.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the text into a Word document using a memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
        {
            Document doc = new Document(stream, loadOptions);

            // Verify that the lists were detected (optional).
            Console.WriteLine($"Detected {doc.Lists.Count} list(s).");

            // Save the resulting document.
            doc.Save("Output.docx");
        }
    }
}
