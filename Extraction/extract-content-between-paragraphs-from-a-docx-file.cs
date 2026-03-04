using System;
using System.IO;
using System.Text;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Example: extract text from paragraph index 2 to 5 (inclusive).
        int startIndex = 2; // zero‑based index of the first paragraph to include.
        int endIndex = 5;   // zero‑based index of the last paragraph to include.

        // Get the collection of paragraphs in the main body.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Clamp indices to valid range.
        if (startIndex < 0) startIndex = 0;
        if (endIndex >= paragraphs.Count) endIndex = paragraphs.Count - 1;

        // Build the extracted text.
        StringBuilder extracted = new StringBuilder();
        for (int i = startIndex; i <= endIndex; i++)
        {
            // GetText returns the paragraph text plus a paragraph break; trim the break.
            extracted.Append(paragraphs[i].GetText().TrimEnd('\r', '\n'));

            // Preserve line breaks between paragraphs.
            if (i < endIndex)
                extracted.AppendLine();
        }

        // Save the extracted content to a plain‑text file.
        File.WriteAllText("Extracted.txt", extracted.ToString());

        // Optional: show full document text using PlainTextDocument.
        PlainTextDocument plain = new PlainTextDocument("Input.docx");
        Console.WriteLine("Full document text:");
        Console.WriteLine(plain.Text);
    }
}
