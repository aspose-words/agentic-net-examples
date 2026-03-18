using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("input.docx");

        // Get the collection of bookmarks from the document's range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Iterate through each bookmark using the enumerator provided by BookmarkCollection.
        using (IEnumerator<Bookmark> enumerator = bookmarks.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                Bookmark bookmark = enumerator.Current;
                if (bookmark != null)
                {
                    // Output the name of the current bookmark.
                    Console.WriteLine(bookmark.Name);
                }
            }
        }

        // Save the document if any modifications were made.
        doc.Save("output.docx");
    }
}
