using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace BookmarkWarningExample
{
    // Custom visitor that checks each BookmarkStart node.
    class BookmarkNameValidator : DocumentVisitor
    {
        public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
        {
            // Log a warning if the bookmark name is null or empty.
            if (string.IsNullOrEmpty(bookmarkStart.Name))
            {
                Console.WriteLine($"Warning: Bookmark at position {bookmarkStart.GetText()} has an empty name.");
            }

            return VisitorAction.Continue;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing document.
            Document doc = new Document("Input.docx");

            // Validate bookmark names.
            BookmarkNameValidator validator = new BookmarkNameValidator();
            doc.Accept(validator);

            // Save the document (unchanged).
            doc.Save("Output.docx");
        }
    }
}
