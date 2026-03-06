using System;
using System.IO;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Path to the input TXT document.
        string inputPath = @"C:\Docs\Input.txt";

        // Load the TXT document. Aspose.Words automatically detects the format.
        Document doc = new Document(inputPath);

        // Get the full text of the document, including control characters.
        // Using the Range.Text property as per the documentation.
        string fullText = doc.Range.Text;

        // Define the start and end markers that delimit the desired content.
        const string startMarker = "[START]";
        const string endMarker = "[END]";

        // Find the positions of the markers.
        int startIndex = fullText.IndexOf(startMarker, StringComparison.Ordinal);
        int endIndex = fullText.IndexOf(endMarker, StringComparison.Ordinal);

        // Validate that both markers were found and are in the correct order.
        if (startIndex == -1 || endIndex == -1 || endIndex <= startIndex)
        {
            Console.WriteLine("Markers not found or are in an invalid order.");
            return;
        }

        // Calculate the start of the content after the start marker.
        int contentStart = startIndex + startMarker.Length;

        // Extract the substring between the two markers.
        string betweenContent = fullText.Substring(contentStart, endIndex - contentStart);

        // Optionally trim any leading/trailing control characters.
        betweenContent = betweenContent.Trim();

        // Output the extracted content.
        Console.WriteLine("Extracted content between markers:");
        Console.WriteLine(betweenContent);
    }
}
