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

        // Write a long sentence that will wrap onto multiple lines.
        string longSentence = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";
        builder.Writeln(longSentence);

        // Update word count and line count properties.
        doc.UpdateWordCount(true);

        // Retrieve the estimated number of lines occupied by the paragraph.
        int lineCount = doc.BuiltInDocumentProperties.Lines;

        // Output the line count.
        Console.WriteLine($"The paragraph occupies {lineCount} line(s).");

        // Save the document to a file.
        doc.Save("ParagraphLines.docx");
    }
}
