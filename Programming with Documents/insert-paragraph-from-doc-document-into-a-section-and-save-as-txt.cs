using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsTxt
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be copied.
        Document srcDoc = new Document("SourceDocument.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Ensure the destination document has a section and a body to hold content.
        // The blank document already contains one section with a body, but we add an explicit section
        // to demonstrate inserting into a specific section if needed.
        Section targetSection = new Section(dstDoc);
        dstDoc.AppendChild(targetSection);
        Body targetBody = new Body(dstDoc);
        targetSection.AppendChild(targetBody);

        // Get the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document using NodeImporter.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the target body (i.e., into the chosen section).
        targetBody.AppendChild(importedParagraph);

        // Save the resulting document as plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions(); // default options; can customize if needed
        dstDoc.Save("Result.txt", txtOptions);
    }
}
