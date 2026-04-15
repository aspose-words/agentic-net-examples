using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph that contains enough text to potentially wrap onto multiple lines when rendered.");

        // Get the paragraph that was just created.
        Paragraph paragraph = builder.CurrentParagraph;

        // Approximate the number of lines in the paragraph.
        // The GetText method returns the paragraph text followed by a paragraph break character (\r).
        // Splitting on this character gives a simple line count approximation.
        string paragraphText = paragraph.GetText();
        int lineCount = paragraphText.Split('\r').Length;

        // Store the result (here we also output it to the console for demonstration).
        Console.WriteLine($"Approximate line count in the paragraph: {lineCount}");

        // Save the document to the file system.
        doc.Save("ParagraphLineCount.docx");
    }
}
