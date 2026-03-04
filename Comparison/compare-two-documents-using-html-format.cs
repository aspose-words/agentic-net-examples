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
            // Perform the comparison. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Configure HTML save options (XHTML 1.0 Transitional with pretty formatting).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            HtmlVersion   = HtmlVersion.Xhtml,
            PrettyFormat  = true
        };

        // Save the compared document as an HTML file.
        docOriginal.Save("ComparisonResult.html", htmlOptions);
    }
}
