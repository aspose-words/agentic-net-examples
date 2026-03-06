using System;
using System.IO;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOTX template.
        Document srcDoc = new Document("Template.dotx");

        // Define the keyword to search for.
        string keyword = "Aspose";

        // Filter paragraphs that contain the keyword (case‑insensitive).
        var matchingParagraphs = srcDoc.FirstSection.Body.Paragraphs
            .Where(p => p.GetText().Contains(keyword, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Create a new blank document to hold the filtered paragraphs.
        Document resultDoc = new Document();

        // Remove the default minimal nodes and create a clean section/body.
        resultDoc.RemoveAllChildren();
        Section section = new Section(resultDoc);
        resultDoc.AppendChild(section);
        Body body = new Body(resultDoc);
        section.AppendChild(body);

        // Import each matching paragraph into the new document, preserving formatting.
        NodeImporter importer = new NodeImporter(srcDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Paragraph para in matchingParagraphs)
        {
            Paragraph imported = (Paragraph)importer.ImportNode(para, true);
            body.AppendChild(imported);
        }

        // Save the filtered paragraphs to a new DOCX file.
        resultDoc.Save("FilteredParagraphs.docx");
    }
}
