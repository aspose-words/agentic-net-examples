using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a long sentence into a single paragraph.
        string longSentence = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                              "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. " +
                              "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.";
        builder.Writeln(longSentence);

        // Recalculate word count and line count.
        doc.UpdateWordCount(true);

        // Get the estimated number of lines occupied by the paragraph.
        int lineCount = doc.BuiltInDocumentProperties.Lines;

        // Output the line count.
        Console.WriteLine($"The paragraph occupies {lineCount} line(s).");

        // Save the document (optional, demonstrates the result).
        doc.Save("ParagraphLines.docx");
    }
}
