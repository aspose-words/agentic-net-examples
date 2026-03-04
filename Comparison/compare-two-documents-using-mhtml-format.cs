using System;
using Aspose.Words;
using Aspose.Words.Saving;

class CompareDocumentsToMhtml
{
    static void Main()
    {
        // Paths to the source documents.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";

        // Load the two documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the comparison result as an MHTML (web archive) file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources to improve compatibility with some mail agents/browsers.
            ExportCidUrlsForMhtmlResources = true,
            // Optional: make the output more readable.
            PrettyFormat = true
        };

        string outputPath = "ComparisonResult.mht";
        docOriginal.Save(outputPath, saveOptions);
    }
}
