using System;
using Aspose.Words;

class TrackChangesDemo
{
    static void Main()
    {
        // Load a DOCX file that contains tracked changes.
        Document doc = new Document("Revisions.docx");

        // -----------------------------------------------------------------
        // 1. Iterate over revision groups.
        //    A group represents a sequence of adjacent revisions made by the same author.
        // -----------------------------------------------------------------
        Console.WriteLine($"Revision groups count: {doc.Revisions.Groups.Count}");
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            Console.WriteLine($"Group author: {group.Author}");
            Console.WriteLine($"Group type  : {group.RevisionType}");
            Console.WriteLine($"Group text  : {group.Text.Trim()}");
            Console.WriteLine();
        }

        // -----------------------------------------------------------------
        // 2. Iterate over individual revisions.
        //    Each Revision object gives detailed information about a single change.
        // -----------------------------------------------------------------
        Console.WriteLine($"Total revisions: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Revision author: {rev.Author}");
            Console.WriteLine($"Revision type  : {rev.RevisionType}");

            // For node‑based revisions (Insertion, Deletion, Moving, FormatChange) the
            // affected text can be obtained from the ParentNode.
            if (rev.RevisionType != RevisionType.StyleDefinitionChange && rev.ParentNode != null)
            {
                Console.WriteLine($"Revision text  : {rev.ParentNode.GetText().Trim()}");
            }
            // For style definition changes the affected style is available via ParentStyle.
            else if (rev.ParentStyle != null)
            {
                Console.WriteLine($"Affected style  : {rev.ParentStyle.Name}");
            }

            Console.WriteLine();
        }

        // -----------------------------------------------------------------
        // 3. (Optional) Accept or reject revisions based on criteria.
        //    Example: accept only insertions made by "John Doe".
        // -----------------------------------------------------------------
        // doc.Revisions.Accept(new RevisionCriteria("John Doe", RevisionType.Insertion));

        // Save the document after processing revisions.
        doc.Save("Revisions_Processed.docx");
    }
}
