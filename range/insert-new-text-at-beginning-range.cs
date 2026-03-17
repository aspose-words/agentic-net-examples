using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add some initial text using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("World!"); // The original content.

        // Retrieve the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Create a Run node that contains the text to be inserted at the beginning.
        Run newRun = new Run(doc, "Hello ");

        // Insert the new Run as the first child of the paragraph.
        firstParagraph.PrependChild(newRun);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
