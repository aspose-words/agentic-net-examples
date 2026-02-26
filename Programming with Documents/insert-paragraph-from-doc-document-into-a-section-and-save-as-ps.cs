using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsPs
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Create a new section and add it to the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);

        // Every section must contain a body. Create it and attach to the section.
        Body body = new Body(dstDoc);
        newSection.AppendChild(body);

        // Import the first paragraph from the source document into the destination document.
        // NodeImporter handles style and list translation between documents.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the new section.
        body.AppendChild(importedParagraph);

        // Save the resulting document as PostScript (PS) using default save options.
        PsSaveOptions psOptions = new PsSaveOptions();
        dstDoc.Save("Result.ps", psOptions);
    }
}
