using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsPs
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.doc");

        // Retrieve the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new (empty) destination document.
        Document dstDoc = new Document();

        // Create a new section and add it to the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);

        // Every section must contain a body; create and attach it.
        Body newBody = new Body(dstDoc);
        newSection.AppendChild(newBody);

        // Import the paragraph from the source document into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the new section's body.
        newBody.AppendChild(importedParagraph);

        // Prepare PostScript save options.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };

        // Save the resulting document as a PostScript file.
        dstDoc.Save("Result.ps", psOptions);
    }
}
