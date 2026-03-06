using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents.
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Compare the documents. Revisions will be added to the original document.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Save the comparison result in DOT format (graph representation of the document structure).
        original.Save("ComparisonResult.dot", SaveFormat.Dot);
    }
}
