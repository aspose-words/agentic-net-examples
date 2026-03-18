using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

class BookmarkConditionalExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example data that could be used to generate a bookmark name.
        // In a real scenario this might come from a database, user input, etc.
        string rawName = GetBookmarkNameFromSomeSource();

        // Evaluate the name expression. If it results in an empty string,
        // we skip creating the bookmark entirely.
        if (!string.IsNullOrEmpty(rawName))
        {
            // Start the bookmark with the evaluated name.
            builder.StartBookmark(rawName);

            // Insert the content that should be inside the bookmark.
            builder.Writeln($"This text is inside bookmark \"{rawName}\".");

            // End the bookmark with the same name.
            builder.EndBookmark(rawName);
        }
        else
        {
            // No bookmark is created; just add regular content.
            builder.Writeln("No bookmark was created because the name expression evaluated to an empty string.");
        }

        // Save the document to disk.
        doc.Save("ConditionalBookmark.docx");
    }

    // Mock method representing the evaluation of a name expression.
    // Returns an empty string to demonstrate the skip logic.
    private static string GetBookmarkNameFromSomeSource()
    {
        // Replace this with actual logic. For demonstration we return empty.
        return string.Empty;
    }
}
