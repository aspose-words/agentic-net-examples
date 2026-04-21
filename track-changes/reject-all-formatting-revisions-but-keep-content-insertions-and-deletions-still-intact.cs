using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial text without tracking – this will not be a revision.
        builder.Writeln("Hello world!");

        // Start tracking revisions.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert new text – this will be an insertion revision.
        builder.Writeln("This is an inserted paragraph.");

        // Change formatting of the first run (make it bold) – this should create a format‑change revision.
        Run firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        firstRun.Font.Bold = true;

        // Delete the word "world" – this will be a deletion revision.
        // The word is part of the first run; split the run to isolate the word.
        string[] parts = firstRun.Text.Split(new[] { "world" }, StringSplitOptions.None);
        if (parts.Length == 2)
        {
            // Replace the original run with two runs, removing the word "world".
            Paragraph para = (Paragraph)firstRun.ParentNode;
            para.InsertBefore(new Run(doc, parts[0]), firstRun);
            para.InsertAfter(new Run(doc, parts[1]), firstRun);
            firstRun.Remove(); // This removal registers as a deletion revision.
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // At this point the document contains insertion, deletion, and format‑change revisions.
        // Reject only the formatting revisions, leaving insertions and deletions intact.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            if (doc.Revisions[i].RevisionType == RevisionType.FormatChange)
                doc.Revisions[i].Reject();
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
