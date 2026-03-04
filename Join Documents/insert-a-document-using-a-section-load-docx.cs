using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank document.
        Document dstDoc = new Document();

        // Remove any default nodes and add a fresh section to the destination document.
        dstDoc.RemoveAllChildren();
        Section dstSection = new Section(dstDoc);
        dstDoc.AppendChild(dstSection);
        // Ensure the new section has a Body and a Paragraph so we can insert content.
        dstSection.EnsureMinimum();

        // Insert the content of the source document's first section at the beginning of the destination section.
        Section srcSection = srcDoc.Sections[0];
        dstSection.PrependContent(srcSection);

        // Save the combined document.
        dstDoc.Save("Result.docx");
    }
}
