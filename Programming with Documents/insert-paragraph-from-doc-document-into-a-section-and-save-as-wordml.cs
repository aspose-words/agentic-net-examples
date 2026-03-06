using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphIntoSection
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.doc");

        // Retrieve the paragraph you want to insert.
        // Here we take the first paragraph of the first section.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new (empty) destination document.
        Document dstDoc = new Document();

        // Create a new section in the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);

        // Ensure the section has a body (and at least one paragraph) so we can add content.
        newSection.EnsureMinimum();

        // Import the source paragraph into the destination document.
        // NodeImporter handles style and list translation between documents.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the new section.
        newSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as WORDML (XML) format.
        dstDoc.Save("Result.xml", SaveFormat.WordML);
    }
}
