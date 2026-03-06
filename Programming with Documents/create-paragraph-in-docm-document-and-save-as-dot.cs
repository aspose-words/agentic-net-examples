using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank DOCM document.
        Document doc = new Document();

        // Ensure the document has a section and a body (already present in a blank document).
        Body body = doc.FirstSection.Body;

        // Create a new paragraph belonging to the document.
        Paragraph paragraph = new Paragraph(doc);
        // Set a style for the paragraph (optional).
        paragraph.ParagraphFormat.StyleName = "Heading 1";

        // Add a run with some text to the paragraph.
        Run run = new Run(doc, "Hello World!");
        paragraph.AppendChild(run);

        // Append the paragraph to the document body.
        body.AppendChild(paragraph);

        // Prepare save options to output the document as a DOT template.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);

        // Save the document as a DOT file.
        doc.Save("OutputTemplate.dot", saveOptions);
    }
}
