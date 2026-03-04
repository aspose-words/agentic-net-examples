using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the DOCM file. The constructor automatically detects the format.
        Document doc = new Document("InputDocument.docm");

        // Assume the region to extract is marked by a bookmark named "ExtractRegion".
        // The bookmark's Text property returns all text between its start and end nodes.
        Bookmark bookmark = doc.Range.Bookmarks["ExtractRegion"];
        if (bookmark == null)
        {
            Console.WriteLine("Bookmark 'ExtractRegion' not found.");
            return;
        }

        // Extract the text contained within the bookmark.
        string extractedText = bookmark.Text;

        // Output the extracted text to the console.
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);

        // Optionally, save the extracted text to a plain‑text file.
        File.WriteAllText("ExtractedContent.txt", extractedText);
    }
}
