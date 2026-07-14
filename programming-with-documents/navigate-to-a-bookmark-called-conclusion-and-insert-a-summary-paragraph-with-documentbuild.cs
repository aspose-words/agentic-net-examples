using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content and a bookmark named "Conclusion".
        builder.Writeln("This is the introduction.");
        builder.StartBookmark("Conclusion");
        builder.Writeln("Conclusion placeholder text.");
        builder.EndBookmark("Conclusion");

        // Move the builder's cursor to the bookmark.
        if (builder.MoveToBookmark("Conclusion"))
        {
            // Insert a summary paragraph at the bookmark location.
            builder.Writeln("Summary: This document demonstrates navigating to a bookmark and inserting a paragraph.");
        }

        // Save the resulting document.
        const string outputFile = "ConclusionExample.docx";
        doc.Save(outputFile);
    }
}
