using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("Introduction paragraph.");
        builder.Writeln("Main content goes here.");

        // Create a bookmark named "Conclusion".
        builder.StartBookmark("Conclusion");
        builder.Writeln("Conclusion placeholder text.");
        builder.EndBookmark("Conclusion");

        // Move the cursor to the position just after the end of the "Conclusion" bookmark.
        // Parameters: bookmarkName, isStart = false (end), isAfter = true (after the end).
        builder.MoveToBookmark("Conclusion", false, true);

        // Insert the summary paragraph.
        builder.Writeln("Summary: This document demonstrates navigating to a bookmark and inserting a paragraph.");

        // Define output path.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DocumentWithSummary.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
