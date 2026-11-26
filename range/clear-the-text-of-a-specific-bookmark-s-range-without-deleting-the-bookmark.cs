using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace BookmarkClearExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and add a bookmark with some text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Paragraph before bookmark.");
            builder.StartBookmark("MyBookmark");
            builder.Writeln("This text is inside the bookmark.");
            builder.EndBookmark("MyBookmark");
            builder.Writeln("Paragraph after bookmark.");

            // Save the original document.
            string originalPath = "Original.docx";
            doc.Save(originalPath);

            // Load the document to demonstrate clearing the bookmark's content.
            Document loadedDoc = new Document(originalPath);

            // Retrieve the bookmark.
            Bookmark bookmark = loadedDoc.Range.Bookmarks["MyBookmark"];
            if (bookmark != null)
            {
                // Remove all nodes that are between the bookmark start and end nodes.
                Node startNode = bookmark.BookmarkStart;
                Node endNode = bookmark.BookmarkEnd;

                Node currentNode = startNode.NextSibling;
                while (currentNode != null && currentNode != endNode)
                {
                    Node nextNode = currentNode.NextSibling;
                    currentNode.Remove();
                    currentNode = nextNode;
                }
            }

            // Save the document after clearing the bookmark's text.
            string clearedPath = "Cleared.docx";
            loadedDoc.Save(clearedPath);
        }
    }
}
