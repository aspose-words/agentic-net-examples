using System;
using System.IO;
using Aspose.Words; // Namespace for PlainTextDocument

class Program
{
    static void Main()
    {
        // Load the plain‑text document from a file.
        string inputPath = "input.txt";
        PlainTextDocument txtDoc = new PlainTextDocument(inputPath);

        // Markers that define the region to extract.
        string startMarker = "[START]";
        string endMarker   = "[END]";

        // Get the full text of the document.
        string fullText = txtDoc.Text;

        // Locate the start and end positions of the region.
        int startIdx = fullText.IndexOf(startMarker, StringComparison.Ordinal);
        int endIdx   = fullText.IndexOf(endMarker, startIdx + startMarker.Length, StringComparison.Ordinal);

        // Extract the substring between the markers, if both are found.
        string extracted = string.Empty;
        if (startIdx != -1 && endIdx != -1 && endIdx > startIdx)
        {
            int contentStart = startIdx + startMarker.Length;
            extracted = fullText.Substring(contentStart, endIdx - contentStart);
        }

        // Save the extracted content to a new text file.
        string outputPath = "extracted.txt";
        File.WriteAllText(outputPath, extracted);
    }
}
