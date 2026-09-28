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
        builder.Writeln("This is the introduction of the document.");
        builder.Writeln();

        // Insert a bookmark named "Conclusion".
        builder.StartBookmark("Conclusion");
        builder.Writeln("Placeholder for the conclusion.");
        builder.EndBookmark("Conclusion");

        // Navigate to the "Conclusion" bookmark and insert a summary paragraph.
        builder.MoveToBookmark("Conclusion");
        builder.Writeln("Summary: This document demonstrates how to navigate to a bookmark and insert text using DocumentBuilder.");

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
