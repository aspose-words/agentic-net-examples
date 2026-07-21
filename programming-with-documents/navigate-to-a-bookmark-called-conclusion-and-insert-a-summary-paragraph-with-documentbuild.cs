using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("This is the beginning of the document.");

        // Insert a bookmark named "Conclusion".
        builder.StartBookmark("Conclusion");
        builder.Writeln("Conclusion placeholder text.");
        builder.EndBookmark("Conclusion");

        // Move the builder's cursor to the start of the "Conclusion" bookmark.
        // The cursor will be positioned just after the bookmark start, inside the bookmark.
        builder.MoveToBookmark("Conclusion");

        // Insert the summary paragraph at the bookmark location.
        builder.Writeln("Summary: This document demonstrates how to navigate to a bookmark and insert text using Aspose.Words.");

        // Save the document to a file in the same directory as the executable.
        doc.Save("Output.docx");
    }
}
