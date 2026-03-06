using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsPs
{
    static void Main()
    {
        // Load the source DOC that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.doc");

        // Create a new empty destination document.
        Document dstDoc = new Document();
        dstDoc.RemoveAllChildren(); // Ensure the document has no default nodes.

        // Create a new section and add it to the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);

        // Every section must contain a body; create and attach it.
        Body body = new Body(dstDoc);
        newSection.AppendChild(body);

        // Prepare a NodeImporter to copy nodes from the source to the destination.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Get the paragraph you want to insert (e.g., the first paragraph of the source).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document.
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the new section.
        body.AppendChild(importedParagraph);

        // Save the resulting document as PostScript (PS) using PsSaveOptions.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };
        dstDoc.Save("Result.ps", psOptions);
    }
}
