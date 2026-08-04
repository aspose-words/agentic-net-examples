using System;
using Aspose.Words;
using Aspose.Words.Properties;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to increase the chance of wrapping.
        builder.PageSetup.PageWidth = 100; // points

        // Write a long sentence into a single paragraph.
        string longSentence = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";
        builder.Write(longSentence);

        // Update word count and line count.
        doc.UpdateWordCount(true);
        int lineCount = doc.BuiltInDocumentProperties.Lines;

        // Display the number of lines occupied by the paragraph.
        Console.WriteLine($"The paragraph occupies {lineCount} line(s).");

        // Save the document to the file system.
        doc.Save("ParagraphLines.docx");
    }
}
