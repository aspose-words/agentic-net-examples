using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some baseline text (not a revision).
        builder.Writeln("Base text.");

        // Track revisions made by Alice.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Alice's addition.");
        doc.StopTrackRevisions();

        // Track revisions made by Bob.
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Bob's addition.");
        doc.StopTrackRevisions();

        // At this point the document contains two insertion revisions.
        Console.WriteLine($"Revisions before reject: {doc.Revisions.Count}");

        // Reject all revisions authored by Bob that are insertions.
        doc.Revisions.Reject(new RevisionCriteria("Bob", RevisionType.Insertion));

        // Confirm that Bob's revision has been removed.
        Console.WriteLine($"Revisions after reject: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Remaining revision author: {rev.Author}");
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }

    // Custom criteria to match a revision by author and type.
    private class RevisionCriteria : IRevisionCriteria
    {
        private readonly string _author;
        private readonly RevisionType _type;

        public RevisionCriteria(string author, RevisionType type)
        {
            _author = author;
            _type = type;
        }

        public bool IsMatch(Revision revision)
        {
            return revision.Author == _author && revision.RevisionType == _type;
        }
    }
}
