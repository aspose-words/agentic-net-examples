using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a paragraph and add a run of text.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "Hello, Aspose.Words! This is a sample paragraph."));

        // Insert the paragraph into the document body.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Copy the paragraph's range content into a string variable.
        string paragraphText = paragraph.Range.Text;

        // Output the extracted text (optional, for verification).
        Console.WriteLine(paragraphText);
    }
}
