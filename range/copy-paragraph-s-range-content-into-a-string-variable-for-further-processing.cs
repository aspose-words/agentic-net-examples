using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a paragraph with sample text.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "This is a sample paragraph."));
        doc.FirstSection.Body.AppendChild(paragraph);

        // Copy the paragraph's range content into a string variable.
        string paragraphContent = paragraph.Range.Text;

        // Example further processing: output the extracted text.
        Console.WriteLine(paragraphContent);
    }
}
