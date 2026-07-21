using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build a paragraph that contains several separate runs (words).
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);
        paragraph.AppendChild(new Run(doc, "Alpha "));
        paragraph.AppendChild(new Run(doc, "Beta "));
        paragraph.AppendChild(new Run(doc, "Gamma "));
        paragraph.AppendChild(new Run(doc, "Delta "));

        // Start tracking revisions – any changes made now will be recorded.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete two consecutive runs ("Beta " and "Gamma ").
        // Remove the later run first to keep indices valid.
        paragraph.Runs[2].Remove(); // Removes "Gamma "
        paragraph.Runs[1].Remove(); // Removes "Beta "

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Find the deletion revision group that contains the two deletions.
        RevisionGroup deletionGroup = null;
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            if (group.RevisionType == RevisionType.Deletion)
            {
                deletionGroup = group;
                break;
            }
        }

        if (deletionGroup == null)
            throw new InvalidOperationException("No deletion revision group was found.");

        // Collect all revisions that belong to the identified group.
        List<Revision> revisionsToAccept = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.Group == deletionGroup)
                revisionsToAccept.Add(rev);
        }

        // Accept each revision. Using a separate list avoids modifying the collection while iterating.
        foreach (Revision rev in revisionsToAccept)
            rev.Accept();

        // After accepting, the document should contain no revisions.
        // Save the resulting document.
        doc.Save("MergedDeletions.docx");
    }
}
