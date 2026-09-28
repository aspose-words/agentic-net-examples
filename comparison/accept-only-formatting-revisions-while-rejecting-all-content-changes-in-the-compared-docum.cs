using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");

        // Create the revised document with a formatting change (bold) and a content change.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Font.Bold = true;
        builder2.Write("Hello");
        builder2.Font.Bold = false;
        builder2.Write(" Aspose.");
        builder2.Writeln();

        // Compare the documents.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Count revisions before processing.
        int totalRevisions = original.Revisions.Count;
        int formattingRevisions = 0;
        int contentRevisions = 0;

        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                formattingRevisions++;
            else
                contentRevisions++;
        }

        // Accept only formatting revisions, reject all others.
        // Iterate over a copy of the collection to avoid modifying it during enumeration.
        while (original.Revisions.Count > 0)
        {
            Revision rev = original.Revisions[0];
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();
            else
                rev.Reject();
        }

        // Verify that all revisions have been processed.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("All revisions should be processed.");

        // Save the resulting document.
        original.Save("Result.docx");
    }
}
