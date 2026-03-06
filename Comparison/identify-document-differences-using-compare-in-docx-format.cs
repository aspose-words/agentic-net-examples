using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the original and edited DOCX files.
        const string originalPath = "Original.docx";
        const string editedPath = "Edited.docx";

        // Load the documents using the Document constructor (lifecycle rule).
        Document original = new Document(originalPath);
        Document edited = new Document(editedPath);

        // Ensure both documents have no existing revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions will be added to the original document.
            original.Compare(edited, "Comparer", DateTime.Now);
        }

        // Enumerate and display each revision (the differences identified).
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the original document now containing the revision marks (save rule).
        original.Save("ComparedResult.docx");
    }
}
