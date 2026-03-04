using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document original = new Document("Original.docx");
        Document edited   = new Document("Edited.docx");

        // Documents must not contain revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Perform the comparison; revisions are added to the original document.
            original.Compare(edited, "Comparer", DateTime.Now);
        }

        // Save the original document (now containing revisions) as HTML.
        original.Save("ComparisonResult.html", SaveFormat.Html);
    }
}
