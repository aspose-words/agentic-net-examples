using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // The Compare method requires that both documents have no existing revisions.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. All differences will be recorded as revisions in docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Prepare save options for MHTML output.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for embedded resources (images, CSS, fonts) to improve compatibility.
            ExportCidUrlsForMhtmlResources = true
        };

        // Save the compared document (with revisions) as an MHTML file.
        docOriginal.Save("ComparisonResult.mht", mhtmlOptions);
    }
}
