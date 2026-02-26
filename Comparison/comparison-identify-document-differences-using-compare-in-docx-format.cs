using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparisonDemo
{
    static void Main()
    {
        // Path to the folder that contains the source documents.
        // Adjust this path to point to your actual files.
        string docsPath = @"C:\Docs\";

        // Load the original and the edited documents.
        Document original = new Document(docsPath + "Original.docx");
        Document edited   = new Document(docsPath + "Edited.docx");

        // The Compare method requires that both documents have no existing revisions.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Optional: configure comparison options (e.g., ignore formatting, case changes, etc.).
            CompareOptions compareOptions = new CompareOptions
            {
                // Example: ignore formatting changes.
                // IgnoreFormatting = true,
                // Example: track changes at the word level (default is WordLevel).
                // Granularity = Granularity.WordLevel
            };

            // Perform the comparison. The revisions are added to the 'original' document.
            original.Compare(edited, "Comparer", DateTime.Now, compareOptions);
        }

        // Output details of each revision that was created.
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document that now contains the tracked changes.
        original.Save(docsPath + "ComparedResult.docx");
    }
}
