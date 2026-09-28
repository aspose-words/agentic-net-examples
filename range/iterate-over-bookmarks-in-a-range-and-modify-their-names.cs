using System;
using Aspose.Words;

namespace BookmarkRenameExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and add some bookmarks.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First bookmark.
            builder.StartBookmark("Bookmark1");
            builder.Writeln("This is the first bookmarked paragraph.");
            builder.EndBookmark("Bookmark1");

            // Second bookmark.
            builder.StartBookmark("Bookmark2");
            builder.Writeln("This is the second bookmarked paragraph.");
            builder.EndBookmark("Bookmark2");

            // Third bookmark.
            builder.StartBookmark("Bookmark3");
            builder.Writeln("This is the third bookmarked paragraph.");
            builder.EndBookmark("Bookmark3");

            // Save the original document.
            const string originalPath = "Original.docx";
            doc.Save(originalPath);

            // Load the document back (demonstrating load workflow).
            Document loadedDoc = new Document(originalPath);

            // Define the range to work with – in this case the whole document.
            Aspose.Words.Range range = loadedDoc.Range;

            // Iterate over all bookmarks within the range and rename them.
            foreach (Bookmark bookmark in range.Bookmarks)
            {
                // Rename the bookmark by prefixing "Renamed_".
                string oldName = bookmark.Name;
                bookmark.Name = "Renamed_" + oldName;
            }

            // Save the modified document.
            const string modifiedPath = "Modified.docx";
            loadedDoc.Save(modifiedPath);

            // Output the new bookmark names to verify the change.
            Console.WriteLine("Renamed bookmarks:");
            foreach (Bookmark bm in loadedDoc.Range.Bookmarks)
            {
                Console.WriteLine(bm.Name);
            }
        }
    }
}
