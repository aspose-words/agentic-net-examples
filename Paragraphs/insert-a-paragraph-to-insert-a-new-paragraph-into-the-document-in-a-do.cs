using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an empty paragraph at the current cursor position.
        // The method returns the newly created Paragraph node.
        Paragraph insertedParagraph = builder.InsertParagraph();

        // Add some text to the newly inserted paragraph.
        // Using Write keeps the cursor inside the same paragraph.
        builder.Write("This is a newly inserted paragraph.");

        // Optionally, insert another paragraph break after the text.
        // This demonstrates that the document now contains two paragraphs.
        builder.InsertParagraph();

        // Save the document in DOCX format.
        doc.Save("InsertedParagraph.docx");
    }
}
