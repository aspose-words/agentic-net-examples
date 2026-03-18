using System;
using Aspose.Words;
// Alias to avoid conflict with System.Range (C# 8+)
using AsposeRange = Aspose.Words.Range;

class Program
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Sample.docx");

        // Get the whole document range.
        AsposeRange fullRange = doc.Range;

        // Retrieve the number of bookmarks in the full range.
        int totalBookmarks = fullRange.Bookmarks.Count;
        Console.WriteLine($"Total bookmarks in the document: {totalBookmarks}");

        // Example: count bookmarks inside a specific sub‑range (first paragraph).
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        AsposeRange subRange = firstParagraph.Range;
        int paragraphBookmarks = subRange.Bookmarks.Count;
        Console.WriteLine($"Bookmarks in the first paragraph: {paragraphBookmarks}");
    }
}
