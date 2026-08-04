using System;
using Aspose.Words;

public class Program
{
    // Criteria that matches revisions belonging to a specific RevisionGroup.
    private class RevisionGroupCriteria : IRevisionCriteria
    {
        private readonly RevisionGroup _targetGroup;

        public RevisionGroupCriteria(RevisionGroup targetGroup)
        {
            _targetGroup = targetGroup;
        }

        public bool IsMatch(Revision revision)
        {
            // Revisions that are part of the target group have their Group property set.
            return revision.Group == _targetGroup;
        }
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that will NOT be tracked.
        builder.Writeln("Original paragraph. ");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert several sequential pieces of text – these will form one RevisionGroup.
        builder.Writeln("First inserted line.");
        builder.Writeln("Second inserted line.");
        builder.Writeln("Third inserted line.");

        // Stop tracking so further edits are not recorded as revisions.
        doc.StopTrackRevisions();

        // Ensure that revisions were created.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("No revisions were generated.");

        // The sequential insertions are grouped; retrieve the first group.
        RevisionGroup firstGroup = doc.Revisions.Groups[0];

        // Accept the entire group with a single call using custom criteria.
        int acceptedCount = doc.Revisions.Accept(new RevisionGroupCriteria(firstGroup));

        // Output information about the operation.
        Console.WriteLine($"Revisions before acceptance: {doc.Revisions.Count + acceptedCount}");
        Console.WriteLine($"Revisions accepted in the group: {acceptedCount}");
        Console.WriteLine($"Revisions remaining after acceptance: {doc.Revisions.Count}");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
