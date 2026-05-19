using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial paragraphs.
        builder.Writeln("Paragraph 1.");
        builder.Writeln("Paragraph 2.");
        builder.Writeln("Paragraph 3.");
        builder.Writeln("Paragraph 4.");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete two consecutive paragraphs to generate two deletion revisions.
        // First deletion.
        doc.FirstSection.Body.Paragraphs[1].Remove(); // Removes "Paragraph 2."
        // Second deletion (the original third paragraph is now at index 1).
        doc.FirstSection.Body.Paragraphs[1].Remove(); // Removes "Paragraph 3."

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // At this point the document contains two consecutive deletion revisions,
        // which are automatically grouped into a single RevisionGroup.
        Console.WriteLine($"Total revision groups: {doc.Revisions.Groups.Count}");

        // Iterate over revision groups to find deletion groups.
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            if (group.RevisionType == RevisionType.Deletion)
            {
                Console.WriteLine($"Found deletion group authored by {group.Author} with text: \"{group.Text.Trim()}\"");

                // Collect revisions belonging to this group before modifying the collection.
                List<Revision> revisionsToAccept = new List<Revision>();
                foreach (Revision rev in doc.Revisions)
                {
                    if (rev.Group == group)
                        revisionsToAccept.Add(rev);
                }

                // Accept each revision in the group.
                foreach (Revision rev in revisionsToAccept)
                    rev.Accept();
            }
        }

        // After accepting the grouped deletions, there should be no remaining revisions.
        Console.WriteLine($"Revisions remaining after acceptance: {doc.Revisions.Count}");

        // Save the resulting document.
        doc.Save("RevisionGroupMerged.docx");
    }
}
