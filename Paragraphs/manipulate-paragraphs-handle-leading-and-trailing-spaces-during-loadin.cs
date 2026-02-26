using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Sample plain‑text with various leading and trailing spaces.
        string txt = "   First line with leading and trailing spaces   \n" +
                     "\tSecond line\t\n" +
                     "Third line    ";

        // Load the text into a memory stream (UTF‑8 encoding).
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(txt)))
        {
            // Configure how leading and trailing spaces are handled during import.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                // Preserve leading spaces so they become part of the paragraph text.
                LeadingSpacesOptions = TxtLeadingSpacesOptions.Preserve,
                // Trim trailing spaces to remove them from the end of each line.
                TrailingSpacesOptions = TxtTrailingSpacesOptions.Trim
            };

            // Load the TXT content into a Word document using the stream and load options.
            Document doc = new Document(stream, loadOptions);

            // Optional: output each paragraph to verify space handling.
            for (int i = 0; i < doc.FirstSection.Body.Paragraphs.Count; i++)
            {
                Paragraph para = (Paragraph)doc.FirstSection.Body.Paragraphs[i];
                // GetText includes paragraph/section marks; TrimEnd removes them for display.
                Console.WriteLine($"Paragraph {i + 1}: \"{para.GetText().TrimEnd('\r', '\n', '\f')}\"");
                // When LeadingSpacesOptions is Preserve, FirstLineIndent remains 0.
                Console.WriteLine($"FirstLineIndent: {para.ParagraphFormat.FirstLineIndent}");
            }

            // Save the resulting document as DOCX.
            doc.Save("Result.docx");
        }
    }
}
