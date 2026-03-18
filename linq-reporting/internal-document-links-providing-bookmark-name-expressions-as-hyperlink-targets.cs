using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Create a bookmark that will be the target of the hyperlink.
        // ------------------------------------------------------------
        const string bookmarkName = "TargetBookmark";

        // Mark the start of the bookmark.
        builder.StartBookmark(bookmarkName);
        // Insert the text that belongs to the bookmark.
        builder.Write("This is the bookmarked text.");
        // Mark the end of the bookmark.
        builder.EndBookmark(bookmarkName);

        // Add a paragraph break after the bookmark so the hyperlink appears on a new line.
        builder.Writeln();

        // ------------------------------------------------------------
        // 2. Insert a hyperlink that points to the bookmark created above.
        // ------------------------------------------------------------
        // Set the visual style for the hyperlink (blue and underlined).
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // InsertHyperlink(displayText, urlOrBookmark, isBookmark)
        // The third argument is true because we are linking to a bookmark, not an external URL.
        Field hyperlink = builder.InsertHyperlink("Jump to bookmark", bookmarkName, true);

        // Optionally set a ScreenTip that appears when the user hovers over the link.
        if (hyperlink is FieldHyperlink fieldHyperlink)
        {
            fieldHyperlink.ScreenTip = "Click to navigate to the bookmarked text.";
        }

        // Reset formatting after the hyperlink if further content is added.
        builder.Font.ClearFormatting();

        // ------------------------------------------------------------
        // 3. Save the document to disk.
        // ------------------------------------------------------------
        const string outputPath = "InternalBookmarkLink.docx";
        doc.Save(outputPath);
    }
}
