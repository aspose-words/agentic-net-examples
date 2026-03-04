using System;
using System.IO;
using System.Text;
using Aspose.Words;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Path to the source HTML file.
        string htmlPath = @"C:\Temp\source.html";

        // Load the HTML document.
        Document doc = new Document(htmlPath);

        // Access the collection of paragraphs in the main body.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Define markers that identify the start and end paragraphs.
        // Adjust these strings to match the exact text of the marker paragraphs.
        const string startMarker = "START";
        const string endMarker = "END";

        int startIndex = -1;
        int endIndex = -1;

        // Locate the indices of the start and end marker paragraphs.
        for (int i = 0; i < paragraphs.Count; i++)
        {
            string text = paragraphs[i].GetText().Trim(); // GetText includes the paragraph break; Trim removes it.
            if (startIndex == -1 && text.Equals(startMarker, StringComparison.OrdinalIgnoreCase))
                startIndex = i;
            else if (startIndex != -1 && text.Equals(endMarker, StringComparison.OrdinalIgnoreCase))
            {
                endIndex = i;
                break;
            }
        }

        // Validate that both markers were found and that there is content between them.
        if (startIndex == -1 || endIndex == -1 || endIndex <= startIndex + 1)
        {
            Console.WriteLine("Unable to locate valid start/end markers or no content between them.");
            return;
        }

        // Extract the text of all paragraphs that lie between the markers.
        StringBuilder extracted = new StringBuilder();
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            // GetText returns the paragraph text plus a trailing paragraph break.
            // TrimEnd removes the trailing break while preserving inner spacing.
            extracted.AppendLine(paragraphs[i].GetText().TrimEnd('\r', '\n'));
        }

        // Output the extracted content to a plain‑text file.
        string outputPath = @"C:\Temp\extracted.txt";
        File.WriteAllText(outputPath, extracted.ToString(), Encoding.UTF8);

        Console.WriteLine($"Content extracted to: {outputPath}");
    }
}
