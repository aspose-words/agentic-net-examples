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
        builder.Writeln("This is a sample paragraph for range extraction.");

        // Retrieve the first paragraph in the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Copy the paragraph's range content into a string variable.
        // The Range.Text property returns the text covered by the paragraph,
        // including the paragraph break character. Trim to remove trailing whitespace.
        string paragraphContent = paragraph.Range.Text.Trim();

        // Example of further processing: convert the extracted text to upper case.
        string processedContent = paragraphContent.ToUpper();

        // Output the processed content to the console.
        Console.WriteLine(processedContent);

        // Save the document to demonstrate the complete lifecycle.
        doc.Save("ParagraphRangeCopy.docx");
    }
}
