using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsMarkdown
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("SourceDocument.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section.
        // (A new Document already contains a default section, body, and paragraph.)
        Section targetSection = dstDoc.FirstSection;

        // Get the first paragraph from the source document.
        // Adjust the index if you need a different paragraph.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph node from the source document into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph at the end of the target section's body.
        // Remove the original empty paragraph that was created by the blank document.
        targetSection.Body.LastParagraph.Remove();
        targetSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as Markdown.
        dstDoc.Save("ResultDocument.md", SaveFormat.Markdown);
    }
}
