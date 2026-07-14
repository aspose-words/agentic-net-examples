using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Plain‑text content that contains list items with different delimiters.
        const string textDoc =
            "Numbered list with dots:\n" +
            "1. First item\n" +
            "2. Second item\n" +
            "3. Third item\n\n" +
            "Numbered list with whitespace delimiter (detectable when the option is true):\n" +
            "1 First item\n" +
            "2 Second item\n" +
            "3 Third item\n\n" +
            "Bulleted list:\n" +
            "• Bullet one\n" +
            "• Bullet two\n" +
            "• Bullet three";

        // Configure loading options to recognise list items that use whitespace as a delimiter.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the plain‑text into a Word document using the configured options.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(textDoc)))
        {
            Document doc = new Document(stream, loadOptions);

            // Ensure list labels are up‑to‑date (optional but recommended).
            doc.UpdateListLabels();

            // Save the resulting document; list items are now proper Word list structures.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedLists.docx");
            doc.Save(outputPath);
        }
    }
}
