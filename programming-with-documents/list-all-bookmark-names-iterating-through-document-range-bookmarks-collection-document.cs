using Aspose.Words;
using System;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path)
        Document doc = new Document("input.docx");

        // Access the collection of bookmarks in the document's range
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Iterate through the collection and print each bookmark's name
        using (IEnumerator<Bookmark> enumerator = bookmarks.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                Bookmark bookmark = enumerator.Current;
                if (bookmark != null)
                {
                    Console.WriteLine(bookmark.Name);
                }
            }
        }

        // Save the document (optional, if you made changes)
        doc.Save("output.docx");
    }
}
