using System;
using Aspose.Words;

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

        // Recalculate word count and line count.
        doc.UpdateWordCount(true);

        // Retrieve the estimated number of lines occupied by the paragraph.
        int lineCount = doc.BuiltInDocumentProperties.Lines;

        // Save the document (optional, demonstrates the lifecycle).
        doc.Save("ParagraphLines.docx");

        // Output the line count.
        Console.WriteLine($"The paragraph occupies {lineCount} line(s).");
    }
}
