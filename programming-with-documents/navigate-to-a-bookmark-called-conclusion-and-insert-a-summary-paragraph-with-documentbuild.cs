using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("This is the beginning of the document.");

        // Insert a bookmark named "Conclusion".
        builder.StartBookmark("Conclusion");
        builder.Writeln("Conclusion placeholder text.");
        builder.EndBookmark("Conclusion");

        // Move the cursor to the end of the "Conclusion" bookmark.
        // Parameters: bookmark name, isStart = false (move to end), isAfter = true (position after the bookmark).
        builder.MoveToBookmark("Conclusion", false, true);

        // Insert the summary paragraph at the bookmark location.
        builder.Writeln("Summary: This paragraph provides a concise summary of the document's conclusions.");

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DocumentWithConclusion.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
