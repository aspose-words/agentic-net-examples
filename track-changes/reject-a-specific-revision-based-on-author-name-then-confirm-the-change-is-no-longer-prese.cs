using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Write("Original text. ");

        // First author makes a revision.
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        builder.Write("John's addition. ");
        doc.StopTrackRevisions();

        // Second author makes a revision.
        doc.StartTrackRevisions("Jane Smith", DateTime.Now);
        builder.Write("Jane's addition. ");
        doc.StopTrackRevisions();

        // Verify that both revisions exist.
        if (doc.Revisions.Count != 2)
            throw new InvalidOperationException("Expected 2 revisions before rejection.");

        // Reject revisions authored by "Jane Smith".
        // Collect matching revisions first to avoid modifying the collection while iterating.
        var revisionsToReject = new System.Collections.Generic.List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.Author == "Jane Smith")
                revisionsToReject.Add(rev);
        }

        foreach (Revision rev in revisionsToReject)
            rev.Reject();

        // Confirm the revision by Jane has been removed.
        if (doc.Revisions.Count != 1)
            throw new InvalidOperationException("Expected 1 revision after rejection.");

        // The document text should no longer contain Jane's addition.
        string expectedText = "Original text. John's addition. ";
        string actualText = doc.GetText();
        if (!actualText.Contains("Jane's addition"))
        {
            // Success: Jane's text is not present.
        }
        else
        {
            throw new InvalidOperationException("Jane's revision was not successfully rejected.");
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
