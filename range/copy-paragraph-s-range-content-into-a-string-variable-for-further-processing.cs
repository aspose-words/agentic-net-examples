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
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This is the second paragraph.");

        // Retrieve the first paragraph in the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Copy the paragraph's range content into a string variable.
        string paragraphText = firstParagraph.Range.Text;

        // The extracted text includes the paragraph break character.
        // Trim it if only the visible text is needed.
        string trimmedText = paragraphText.Trim();

        // Example of further processing: display the extracted text.
        Console.WriteLine("Extracted paragraph text:");
        Console.WriteLine(trimmedText);
    }
}
