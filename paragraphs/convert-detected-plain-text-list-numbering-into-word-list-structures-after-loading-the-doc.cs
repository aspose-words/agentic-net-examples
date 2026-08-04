using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample plain‑text that contains numbered list items.
        const string plainText = 
            "1. First item\n" +
            "2. Second item\n" +
            "3. Third item\n\n" +
            "1 Fourth item\n" +
            "2 Fourth item\n" +
            "3 Fourth item";

        // Configure loading options to detect list numbering, including whitespace delimiters.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };

        // Load the plain‑text into a Word document using the configured options.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(plainText)))
        {
            Document doc = new Document(stream, loadOptions);

            // Ensure that list labels are up‑to‑date after loading.
            doc.UpdateListLabels();

            // Save the resulting document with proper Word list structures.
            doc.Save("ConvertedList.docx");
        }
    }
}
