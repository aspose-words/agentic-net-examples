using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original text.");

        // Insert new text at the beginning of the document's range.
        // Use the first paragraph as the reference node.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Create a new paragraph that contains the inserted text.
        Paragraph insertedParagraph = new Paragraph(doc);
        insertedParagraph.AppendChild(new Run(doc, "Inserted text. "));

        // Insert the new paragraph before the first existing paragraph.
        firstParagraph.ParentNode.InsertBefore(insertedParagraph, firstParagraph);

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Output the resulting document text to the console for verification.
        Console.WriteLine("Document text after insertion:");
        Console.WriteLine(doc.Range.Text);
    }
}
