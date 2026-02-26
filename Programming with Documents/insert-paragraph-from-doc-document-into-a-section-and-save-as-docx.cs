using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC file.
        Document srcDoc = new Document("Source.doc");

        // Retrieve the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new blank destination document.
        Document dstDoc = new Document();
        dstDoc.EnsureMinimum(); // Guarantees at least one section and body.

        // Get the target section where the paragraph will be inserted.
        Section targetSection = dstDoc.FirstSection;

        // Import the paragraph node from the source document into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the target section.
        targetSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as DOCX.
        dstDoc.Save("Result.docx");
    }
}
