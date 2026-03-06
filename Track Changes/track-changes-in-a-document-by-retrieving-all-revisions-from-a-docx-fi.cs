using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file that contains tracked changes.
        string inputPath = "TrackedChanges.docx";

        // Load the document from the file system (load rule).
        Document doc = new Document(inputPath);

        // Output whether the document contains any revisions.
        Console.WriteLine($"Document has revisions: {doc.HasRevisions}");

        // Iterate through all revisions in the document (retrieving tracked changes).
        foreach (Revision revision in doc.Revisions)
        {
            // Author who made the change.
            Console.WriteLine($"Author: {revision.Author}");

            // Date and time when the revision was created.
            Console.WriteLine($"Date: {revision.DateTime}");

            // Type of revision (Insertion, Deletion, etc.).
            Console.WriteLine($"Type: {revision.RevisionType}");

            // Text associated with the revision. Trim to remove extra whitespace.
            Console.WriteLine($"Text: {revision.ParentNode.GetText().Trim()}");

            Console.WriteLine(new string('-', 40));
        }

        // Optionally, save a copy of the document (save rule).
        string outputPath = "TrackedChanges_Copy.docx";
        doc.Save(outputPath);
    }
}
