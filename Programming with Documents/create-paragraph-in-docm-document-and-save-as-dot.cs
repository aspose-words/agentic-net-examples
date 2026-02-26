using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document (default minimal structure).
        Document doc = new Document();

        // Add a paragraph with some text.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.ParagraphFormat.StyleName = "Normal";
        paragraph.AppendChild(new Run(doc, "This is a paragraph in a DOCM document."));
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document as a macro‑enabled DOCM file.
        string docmPath = "Document.docm";
        doc.Save(docmPath, SaveFormat.Docm);

        // Load the previously saved DOCM document.
        Document loadedDoc = new Document(docmPath);

        // Prepare save options to output a DOT (template) file.
        DocSaveOptions dotOptions = new DocSaveOptions(SaveFormat.Dot);

        // Save the loaded document as a DOT template.
        string dotPath = "Document.dot";
        loadedDoc.Save(dotPath, dotOptions);
    }
}
