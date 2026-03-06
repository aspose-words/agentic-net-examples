using System;
using Aspose.Words;

class InsertDocumentAtBookmark
{
    static void Main()
    {
        // Paths to the documents.
        string destinationPath = "Destination.docx";   // Document that contains the bookmark.
        string sourcePath = "Source.docx";            // Document to be inserted.
        string outputPath = "Result.docx";            // Where the merged document will be saved.

        // Load the destination document (which must already contain a bookmark named "InsertHere").
        Document destDoc = new Document(destinationPath);

        // Load the source document that will be inserted.
        Document srcDoc = new Document(sourcePath);

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destDoc);

        // Move the builder's cursor to the bookmark.
        // The bookmark name must match the one defined in the destination document.
        bool bookmarkFound = builder.MoveToBookmark("InsertHere");
        if (!bookmarkFound)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found in the destination document.");

        // Insert the source document at the bookmark position.
        // Use the desired import format mode (e.g., keep source formatting or use destination styles).
        builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // Save the resulting document.
        destDoc.Save(outputPath);
    }
}
