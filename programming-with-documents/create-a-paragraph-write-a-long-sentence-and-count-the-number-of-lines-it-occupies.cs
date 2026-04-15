using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a long sentence.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. " +
                        "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. " +
                        "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.");

        // Update word count and line count properties.
        doc.UpdateWordCount(true);
        int lineCount = doc.BuiltInDocumentProperties.Lines;

        // Output the number of lines the paragraph occupies.
        Console.WriteLine($"The paragraph occupies {lineCount} line(s).");

        // Save the document to verify the content (optional but follows core rules).
        doc.Save("output.docx");
    }
}
