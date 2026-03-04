using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document (DOCM format is supported by default).
        Document doc = new Document();

        // Create a new paragraph belonging to the document.
        Paragraph paragraph = new Paragraph(doc);

        // Set paragraph formatting (optional).
        paragraph.ParagraphFormat.StyleName = "Heading 1";
        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Add a run with some text to the paragraph.
        Run run = new Run(doc, "Hello from DOCM!");
        paragraph.AppendChild(run);

        // Append the paragraph to the body of the first section.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Prepare save options to output a DOT (Word template) file.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);

        // Save the document as a DOT file.
        doc.Save("Output.dot", saveOptions);
    }
}
