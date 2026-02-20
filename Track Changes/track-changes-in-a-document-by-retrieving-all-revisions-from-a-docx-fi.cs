using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains tracked changes.
        Document doc = new Document("Revisions.docx");

        // Get the collection of revisions from the document.
        RevisionCollection revisions = doc.Revisions;

        // Display the total number of revisions found.
        Console.WriteLine($"Total revisions: {revisions.Count}");

        // Iterate through each revision and output its details.
        foreach (Revision rev in revisions)
        {
            // For style changes the affected text is stored in ParentStyle,
            // otherwise it is stored in ParentNode.
            string affectedText = rev.RevisionType == RevisionType.StyleDefinitionChange
                ? $"Style: {rev.ParentStyle?.Name}"
                : rev.ParentNode?.GetText().Trim();

            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime}");
            Console.WriteLine($"Type: {rev.RevisionType}");
            Console.WriteLine($"Affected Text: {affectedText}");
            Console.WriteLine(new string('-', 40));
        }

        // Optionally, display information about revision groups (condensed view).
        Console.WriteLine($"Revision groups: {revisions.Groups.Count}");
        foreach (RevisionGroup group in revisions.Groups)
        {
            Console.WriteLine($"Group Author: {group.Author}");
            Console.WriteLine($"Group Type: {group.RevisionType}");
            Console.WriteLine($"Group Text: {group.Text.Trim()}");
            Console.WriteLine(new string('=', 40));
        }
    }
}
