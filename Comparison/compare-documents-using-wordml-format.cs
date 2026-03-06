using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareWordmlExample
{
    static void Main()
    {
        // Load the original document.
        Document original = new Document("Original.docx");

        // Load the document to compare against.
        Document revised = new Document("Revised.docx");

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (original.Revisions.Count == 0 && revised.Revisions.Count == 0)
        {
            // Perform the comparison. All differences will be stored as revisions in the original document.
            original.Compare(revised, "Comparer", DateTime.Now);
        }

        // Save the comparison result in WORDML (XML) format.
        // The .xml extension tells Aspose.Words to use the WORDML save format.
        original.Save("ComparisonResult.xml");
    }
}
