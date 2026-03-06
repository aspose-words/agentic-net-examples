using Aspose.Words;
using System;

class InsertParagraphExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first paragraph.
        builder.Writeln("This is the first paragraph.");

        // Insert an empty paragraph break at the current cursor position.
        // The method returns the newly created Paragraph node.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Write a second paragraph after the inserted empty paragraph.
        builder.Writeln("This is the second paragraph after an empty one.");

        // Save the document to a DOCX file.
        doc.Save("InsertedParagraph.docx");
    }
}
