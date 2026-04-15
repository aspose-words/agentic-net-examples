using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a couple of paragraphs.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This is the second paragraph.");

        // Retrieve the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Copy the paragraph's range content into a string variable.
        string paragraphText = firstParagraph.Range.Text;

        // Output the extracted text (trimmed to remove trailing control characters).
        Console.WriteLine(paragraphText.Trim());

        // Save the document (optional, demonstrates saving workflow).
        doc.Save("Output.docx");
    }
}
