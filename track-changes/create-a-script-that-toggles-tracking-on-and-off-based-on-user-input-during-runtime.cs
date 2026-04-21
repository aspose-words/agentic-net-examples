using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content without tracking.
        builder.Writeln("Initial content. ");

        // Simulated user toggle sequence: true = start tracking, false = stop tracking.
        bool[] toggleSequence = new bool[] { true, false, true };

        for (int i = 0; i < toggleSequence.Length; i++)
        {
            if (toggleSequence[i])
            {
                // Enable tracking with a sample author.
                doc.StartTrackRevisions("User", DateTime.Now);
                builder.Writeln($"Tracked change #{i + 1}. ");
            }
            else
            {
                // Disable tracking.
                doc.StopTrackRevisions();
                builder.Writeln($"Untracked change #{i + 1}. ");
            }
        }

        // Ensure tracking is stopped before final save.
        doc.StopTrackRevisions();

        // Output revision information.
        Console.WriteLine($"Total revisions: {doc.Revisions.Count}");
        for (int i = 0; i < doc.Revisions.Count; i++)
        {
            Revision rev = doc.Revisions[i];
            Console.WriteLine($"Revision {i + 1}: Author={rev.Author}, Type={rev.RevisionType}, Text=\"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document to a local file.
        doc.Save("TrackedDocument.docx");
    }
}
