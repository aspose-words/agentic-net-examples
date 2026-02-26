// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;

class InsertTextAfterRange
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Define the range you want to work with.
        // For demonstration, use the whole document range.
        // In practice, obtain a specific range (e.g., a bookmark, a table cell, etc.).
        Range range = doc.Range;

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the end of the range.
        // The Range.End property returns the node that marks the end of the range.
        builder.MoveTo(range.End);

        // Insert the desired text after the range.
        builder.Write("Inserted text");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
