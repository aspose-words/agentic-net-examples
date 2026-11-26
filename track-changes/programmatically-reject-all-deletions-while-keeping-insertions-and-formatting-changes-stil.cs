using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with two paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph 1.");
        builder.Writeln("Original paragraph 2.");

        // Start tracking revisions.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // 1. Delete the first paragraph (creates a deletion revision).
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.Remove();

        // 2. Insert a new paragraph (creates an insertion revision).
        builder.Writeln("This is an inserted paragraph.");

        // 3. Change formatting of the original second paragraph (creates a format change revision).
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[0]; // after removal, this is the original second paragraph.
        builder.MoveTo(secondParagraph);
        builder.Font.Italic = true;
        builder.Write(" (now italic)");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only deletions, keep insertions and format changes.
        List<Revision> deletions = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Deletion)
                deletions.Add(rev);
        }

        foreach (Revision rev in deletions)
        {
            rev.Reject();
        }

        // Save the resulting document.
        string outputPath = "TrackChangesResult.docx";
        doc.Save(outputPath);

        // Output summary.
        Console.WriteLine($"Revisions processed. Document saved to '{outputPath}'.");
        Console.WriteLine($"Total revisions remaining: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }
}
