using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("Introduction");
        builder.Writeln();

        // Insert a bookmark named "Conclusion".
        builder.StartBookmark("Conclusion");
        builder.Writeln("This is the conclusion placeholder.");
        builder.EndBookmark("Conclusion");

        // Move the builder's cursor to the start of the bookmark.
        if (builder.MoveToBookmark("Conclusion"))
        {
            // Insert the summary paragraph at the bookmark location.
            builder.Writeln("Summary: This document demonstrates navigating to a bookmark and inserting text.");
        }

        // Save the document to the local file system.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
