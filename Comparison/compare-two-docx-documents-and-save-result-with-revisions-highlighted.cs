using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the source documents and the output HTML file.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";
        string resultPath   = "ComparisonResult.html";

        // Load the two documents that will be compared.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions (insertions/deletions) will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the resulting document as HTML.
        // The .html extension automatically selects the HTML save format.
        // Revisions are exported as <ins> and <del> tags in the HTML output.
        docOriginal.Save(resultPath);
    }
}
