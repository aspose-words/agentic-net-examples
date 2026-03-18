using System;
using Aspose.Words;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a long sentence that will wrap onto multiple lines.
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua, " +
            "ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Update word count and calculate the number of lines in the document.
        // The overload with 'true' forces Aspose.Words to recompute the line count.
        doc.UpdateWordCount(true);

        // Retrieve the estimated number of lines from built‑in document properties.
        int lineCount = doc.BuiltInDocumentProperties.Lines;

        // Output the line count.
        Console.WriteLine($"The paragraph occupies {lineCount} line(s).");

        // Save the document (optional, demonstrates the required save lifecycle step).
        doc.Save("ParagraphLineCount.docx");
    }
}
