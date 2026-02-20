using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class MarkdownParagraphExtractor
{
    // Extracts the text that lies between two paragraphs identified by their exact text.
    // startMarker – text of the paragraph that marks the beginning (exclusive).
    // endMarker   – text of the paragraph that marks the end (exclusive).
    public static string ExtractBetween(string markdownPath, string startMarker, string endMarker)
    {
        // Load the markdown file into an Aspose.Words Document.
        var loadOptions = new MarkdownLoadOptions();
        Document doc = new Document(markdownPath, loadOptions);

        // Get the collection of paragraphs in the main story (the document body).
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        int startIndex = -1;
        int endIndex = -1;

        // Locate the start and end markers.
        for (int i = 0; i < paragraphs.Count; i++)
        {
            string text = paragraphs[i].GetText().TrimEnd('\r', '\n');
            if (startIndex == -1 && text == startMarker)
                startIndex = i;
            else if (endIndex == -1 && text == endMarker)
                endIndex = i;
        }

        // Validate that both markers were found and are in the correct order.
        if (startIndex == -1)
            throw new InvalidOperationException($"Start marker \"{startMarker}\" not found.");
        if (endIndex == -1)
            throw new InvalidOperationException($"End marker \"{endMarker}\" not found.");
        if (endIndex <= startIndex)
            throw new InvalidOperationException("End marker appears before start marker.");

        // Concatenate the text of all paragraphs between the markers.
        StringWriter writer = new StringWriter();
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            writer.Write(paragraphs[i].GetText());
        }

        return writer.ToString();
    }

    // Example usage.
    static void Main()
    {
        string mdFile = @"C:\Docs\sample.md";

        // Define the exact paragraph texts that bound the region to extract.
        string start = "## Start Section";
        string end   = "## End Section";

        try
        {
            string extracted = ExtractBetween(mdFile, start, end);
            Console.WriteLine("Extracted content:");
            Console.WriteLine(extracted);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
