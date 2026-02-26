using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains tracked changes.
        // The Document constructor handles the loading operation.
        Document doc = new Document("Revisions.docx");

        // Ensure the document actually has revisions before processing.
        if (!doc.HasRevisions)
        {
            Console.WriteLine("The document does not contain any revisions.");
            return;
        }

        // Iterate through each revision group.
        // A RevisionGroup represents a sequential set of revisions made by the same author.
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            // Output basic information about the group.
            Console.WriteLine($"Author: {group.Author}");
            Console.WriteLine($"Revision Type: {group.RevisionType}");
            Console.WriteLine($"Text: {group.Text}");
            Console.WriteLine(new string('-', 40));
        }

        // Optionally, iterate through individual revisions for more granular details.
        Console.WriteLine("Individual revisions:");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime}");
            Console.WriteLine($"Type: {rev.RevisionType}");
            // The ParentNode contains the node that was changed.
            // GetText() returns the text of that node (or a description for format changes).
            Console.WriteLine($"Changed Text: {rev.ParentNode?.GetText().Trim()}");
            Console.WriteLine(new string('-', 40));
        }

        // Save the document unchanged (optional, demonstrates the save lifecycle rule).
        doc.Save("Revisions_Processed.docx");
    }
}
