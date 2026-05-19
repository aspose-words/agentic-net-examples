using System;
using System.IO;
using Aspose.Words;

namespace BookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some introductory text.
            builder.Writeln("This is the introduction of the document.");

            // Insert a bookmark named "Conclusion" with placeholder text.
            builder.StartBookmark("Conclusion");
            builder.Writeln("Conclusion placeholder text.");
            builder.EndBookmark("Conclusion");

            // Move the cursor to the position just after the end of the "Conclusion" bookmark.
            // Parameters: bookmark name, isStart = false (end of bookmark), isAfter = true (after the end).
            builder.MoveToBookmark("Conclusion", false, true);

            // Insert the summary paragraph at the bookmark location.
            builder.Writeln("Summary: This document demonstrates how to navigate to a bookmark and insert text using Aspose.Words.");

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

            // Save the document to the specified file.
            doc.Save(outputPath);

            // Optional: Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                Console.WriteLine("Failed to save the document.");
            }
        }
    }
}
