using System;
using Aspose.Words;

namespace BookmarkClearExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing document (replace with your actual file path).
            Document doc = new Document("Input.docx");

            // Name of the bookmark whose text we want to clear.
            const string bookmarkName = "MyBookmark";

            // Retrieve the bookmark from the document's bookmark collection.
            Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];

            if (bookmark != null)
            {
                // Clear the text inside the bookmark while preserving the bookmark itself.
                bookmark.Text = string.Empty;
            }
            else
            {
                Console.WriteLine($"Bookmark '{bookmarkName}' not found.");
            }

            // Save the modified document (replace with your desired output path).
            doc.Save("Output.docx");
        }
    }
}
