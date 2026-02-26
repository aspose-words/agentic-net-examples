using System;
using Aspose.Words;
using Aspose.Words.Saving;

class CompareDocsToHtml
{
    static void Main()
    {
        // Load the original and edited DOCX files.
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Perform the comparison. Revisions are added to the original document.
            original.Compare(edited, "Comparer", DateTime.Now);
        }

        // Save the result as HTML; revisions will be highlighted in the output.
        original.Save("ComparisonResult.html", SaveFormat.Html);
    }
}
