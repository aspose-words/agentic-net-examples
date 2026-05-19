using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a couple of paragraphs using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This is the second paragraph.");

        // Retrieve the first paragraph node.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Copy the paragraph's range content into a string variable.
        string paragraphContent = firstParagraph.Range.Text;

        // Output the extracted text (for demonstration purposes).
        Console.WriteLine("Extracted paragraph text:");
        Console.WriteLine(paragraphContent);

        // Save the document to the local file system (optional).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.docx");
        doc.Save(outputPath);
    }
}
