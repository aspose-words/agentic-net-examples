using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a paragraph with initial text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Hello world.");
        paragraph.AppendChild(run);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Append text to the end of the document's range.
        // Get the last paragraph in the body and add a new Run with the desired text.
        Paragraph lastParagraph = doc.FirstSection.Body.LastParagraph;
        lastParagraph.AppendChild(new Run(doc, " Appended text."));

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Output the final document text.
        Console.WriteLine("Document text after appending:");
        Console.WriteLine(doc.GetText());
    }
}
