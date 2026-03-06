using System;
using Aspose.Words;
using Aspose.Words.Saving;

class CompareDocumentsToHtml
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Configure HTML save options (optional: choose HTML5 version).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            HtmlVersion = HtmlVersion.Html5,
            PrettyFormat = true
        };

        // Save the compared document (with revisions) as HTML.
        docOriginal.Save("ComparisonResult.html", htmlOptions);
    }
}
