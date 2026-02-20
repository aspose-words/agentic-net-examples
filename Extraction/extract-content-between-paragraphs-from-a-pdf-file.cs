using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Docs\source.pdf";

        // Load the PDF into an Aspose.Words Document using PdfLoadOptions.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(pdfPath, loadOptions);

        // Get the collection of paragraphs in the main text story.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Define the start and end paragraph indices (inclusive).
        // Adjust these indices as needed for the specific paragraphs you want to target.
        int startIndex = 2; // third paragraph (0‑based index)
        int endIndex = 5;   // sixth paragraph (0‑based index)

        // Validate indices.
        if (startIndex < 0 || endIndex >= paragraphs.Count || startIndex > endIndex)
        {
            Console.WriteLine("Invalid paragraph range.");
            return;
        }

        // Build the text that lies between the selected paragraphs.
        StringBuilder sb = new StringBuilder();
        for (int i = startIndex; i <= endIndex; i++)
        {
            // Paragraph.GetText() returns the paragraph text including its terminating \r.
            sb.Append(paragraphs[i].GetText());
        }

        string extractedText = sb.ToString();

        // Output the extracted text.
        Console.WriteLine("Extracted text between paragraphs:");
        Console.WriteLine(extractedText);
    }
}
