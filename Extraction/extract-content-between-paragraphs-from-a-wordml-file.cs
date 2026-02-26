using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the WORDML (or any supported Word) document.
        // The constructor automatically detects the format, so a .xml WordML file works here.
        Document doc = new Document("Input.docx");

        // Access the collection of paragraphs in the main body of the first section.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Define markers that bound the region we want to extract.
        // In this example we capture everything between a paragraph that contains "Start"
        // and a paragraph that contains "End".
        bool capture = false;
        List<string> extractedLines = new List<string>();

        foreach (Paragraph para in paragraphs)
        {
            // Get the visible text of the paragraph (trimmed to remove the trailing paragraph break).
            string text = para.GetText().Trim();

            if (text.Equals("Start", StringComparison.OrdinalIgnoreCase))
            {
                // Begin capturing after the start marker.
                capture = true;
                continue; // Skip the marker paragraph itself.
            }

            if (text.Equals("End", StringComparison.OrdinalIgnoreCase))
            {
                // Stop capturing when the end marker is reached.
                capture = false;
                break; // Exit the loop if only the first region is needed.
            }

            if (capture)
            {
                // Collect the paragraph text.
                extractedLines.Add(text);
            }
        }

        // Combine the captured paragraphs into a single string.
        string result = string.Join(Environment.NewLine, extractedLines);

        // Save the extracted content to a plain‑text file (optional).
        File.WriteAllText("Extracted.txt", result);
    }
}
