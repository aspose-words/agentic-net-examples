using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Path to the source RTF file.
        const string inputPath = @"C:\Docs\source.rtf";

        // Load the RTF document with default load options.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Retrieve the full text of the document (includes control characters).
        string rawText = doc.GetText();

        // Normalize line breaks for easier processing.
        string normalizedText = rawText.Replace(ControlChar.Cr.ToString(), "\n")
                                      .Replace(ControlChar.Lf.ToString(), "\n");

        // Define the markers that bound the desired content.
        const string startMarker = "[START]";
        const string endMarker   = "[END]";

        // Locate the markers.
        int startIndex = normalizedText.IndexOf(startMarker, StringComparison.Ordinal);
        int endIndex   = normalizedText.IndexOf(endMarker,   StringComparison.Ordinal);

        // Extract the text between the markers, if both are found.
        string extracted = string.Empty;
        if (startIndex != -1 && endIndex != -1 && endIndex > startIndex)
        {
            int contentStart = startIndex + startMarker.Length;
            extracted = normalizedText.Substring(contentStart,
                                                endIndex - contentStart).Trim();
        }

        // Save the extracted content to a plain‑text file.
        const string outputPath = @"C:\Docs\extracted.txt";
        File.WriteAllText(outputPath, extracted);
    }
}
