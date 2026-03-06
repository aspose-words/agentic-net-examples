using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveMhtml
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to be inserted.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Add a new empty section to the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);
        Body body = new Body(dstDoc);
        newSection.AppendChild(body);

        // Position the DocumentBuilder at the beginning of the newly added section.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.MoveToSection(dstDoc.Sections.Count - 1);

        // Insert the content of the source document (its paragraph) into the current position.
        // Keep the source formatting while importing.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document as MHTML (Web archive) format.
        dstDoc.Save("Result.mht", SaveFormat.Mhtml);
    }
}
