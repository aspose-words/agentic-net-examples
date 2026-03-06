using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.docx");

        // Retrieve the paragraph you want to insert (e.g., the first paragraph).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new destination document.
        Document dstDoc = new Document(); // Document constructor creates a default section.

        // Import the paragraph node from the source document into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph at the end of the first section's body.
        Body dstBody = dstDoc.FirstSection.Body;
        if (dstBody.HasChildNodes)
        {
            // Insert after the last existing node in the body.
            Node lastNode = dstBody.LastChild;
            dstDoc.InsertAfter(importedParagraph, lastNode);
        }
        else
        {
            // If the body is empty, simply append the paragraph.
            dstBody.AppendChild(importedParagraph);
        }

        // Save the resulting document as XPS.
        XpsSaveOptions saveOptions = new XpsSaveOptions();
        dstDoc.Save("Result.xps", saveOptions);
    }
}
