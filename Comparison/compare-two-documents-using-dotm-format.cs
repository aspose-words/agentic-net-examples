using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original DOTM document.
        Document original = new Document("Original.dotm");

        // Load the edited DOTM document to compare against.
        Document edited = new Document("Edited.dotm");

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions will be added to the original document.
            original.Compare(edited, "JD", DateTime.Now);
        }

        // Save the comparison result as a DOTM file.
        original.Save("ComparisonResult.dotm");
    }
}
