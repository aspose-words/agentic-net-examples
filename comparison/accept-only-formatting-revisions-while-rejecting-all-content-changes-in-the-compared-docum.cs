using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");                     // Plain text.
        builderOriginal.Font.Bold = true;
        builderOriginal.Writeln("Bold text.");                       // Bold formatting.

        // Create the revised document with both content and formatting changes.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello world changed.");              // Content change.
        builderRevised.Font.Bold = false;                            // Formatting change (remove bold).
        builderRevised.Writeln("Bold text.");                        // Same text, different formatting.
        builderRevised.Writeln("Additional paragraph.");             // Insertion.

        // Perform the comparison. The original document will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were created during comparison.");

        // Accept only formatting revisions; reject all other types.
        // Iterate over a copy because accepting/rejecting modifies the collection.
        List<Revision> revisions = original.Revisions.Cast<Revision>().ToList();
        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();   // Keep formatting changes.
            else
                rev.Reject();   // Discard other changes.
        }

        // After processing, there should be no remaining revisions.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Some revisions were not processed correctly.");

        // Save the resulting document.
        original.Save("Result.docx");
    }
}
