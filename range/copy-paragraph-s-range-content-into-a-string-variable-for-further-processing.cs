using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph whose range will be extracted.");

        // Retrieve the first paragraph in the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Copy the paragraph's range content into a string variable.
        string paragraphContent = paragraph.Range.Text;

        // Output the extracted text to verify the operation.
        Console.WriteLine("Extracted paragraph text:");
        Console.WriteLine(paragraphContent.Trim());

        // Optionally save the document to demonstrate the full lifecycle.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "SampleDocument.docx");
        doc.Save(outputPath);
    }
}
