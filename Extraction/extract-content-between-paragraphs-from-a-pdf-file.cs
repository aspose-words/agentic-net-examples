using System;
using System.IO;
using System.Text;
using Aspose.Words;

class ExtractParagraphContent
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Docs\source.pdf";

        // Load the PDF as an Aspose.Words document.
        Document doc = new Document(pdfPath);

        // Collect the text of each paragraph.
        StringBuilder extractedText = new StringBuilder();

        // Iterate through all paragraphs in the main body of the first section.
        foreach (Paragraph paragraph in doc.FirstSection.Body.Paragraphs)
        {
            // GetText() returns the paragraph text including the paragraph break character.
            // Trim the trailing paragraph break and any surrounding whitespace.
            string text = paragraph.GetText().TrimEnd('\r', '\n').Trim();

            // Append the cleaned text to the result.
            extractedText.AppendLine(text);
        }

        // Save the extracted content to a plain‑text file.
        string outputPath = @"C:\Docs\extracted.txt";
        File.WriteAllText(outputPath, extractedText.ToString());

        Console.WriteLine($"Extraction complete. Output saved to: {outputPath}");
    }
}
