using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsBookmarkDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX template that contains bookmarks.
            const string inputPath = @"C:\Docs\TemplateWithBookmarks.docx";

            // Path where the resulting document will be saved.
            const string outputPath = @"C:\Docs\ResultWithBookmarks.docx";

            // Load the existing document.
            Document doc = new Document(inputPath);

            // Example data to insert into bookmarks.
            // The key is the bookmark name, the value is the text to place at that bookmark.
            var bookmarkData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Title", "Quarterly Report" },
                { "Author", "John Doe" },
                { "Date", DateTime.Today.ToString("MMMM dd, yyyy") },
                { "Summary", "This quarter showed a 15% increase in sales." }
            };

            // Use LINQ to enumerate all bookmarks in the document.
            // Cast the collection to IEnumerable<Bookmark> for LINQ support.
            var bookmarks = doc.Range.Bookmarks.Cast<Bookmark>();

            foreach (Bookmark bm in bookmarks)
            {
                // If we have data for the current bookmark, replace its contents.
                if (bookmarkData.TryGetValue(bm.Name, out string newText))
                {
                    // The Bookmark.Text property replaces the entire content of the bookmark.
                    bm.Text = newText;
                }
            }

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
