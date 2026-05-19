using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Criteria that matches revisions authored by a specific user.
    private class AuthorRevisionCriteria : IRevisionCriteria
    {
        private readonly string _authorName;

        public AuthorRevisionCriteria(string authorName)
        {
            _authorName = authorName ?? throw new ArgumentNullException(nameof(authorName));
        }

        public bool IsMatch(Revision revision)
        {
            return revision != null && revision.Author == _authorName;
        }
    }

    public static void Main()
    {
        // Folder that contains the documents to process.
        string inputFolder = Path.Combine(Environment.CurrentDirectory, "Docs");
        Directory.CreateDirectory(inputFolder);

        // Create sample documents with revisions.
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"));

        // Author whose revisions should be rejected.
        string targetAuthor = "Bob";

        // Process each .docx file in the folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Reject all revisions authored by the target user.
            doc.Revisions.Reject(new AuthorRevisionCriteria(targetAuthor));

            // Overwrite the original file with the cleaned document.
            doc.Save(filePath);
        }
    }

    // Creates a sample document containing revisions from two different authors.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (no revisions).
        builder.Writeln("Original paragraph.");

        // Revisions by Alice.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Alice's inserted paragraph.");
        doc.StopTrackRevisions();

        // Revisions by Bob.
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Bob's inserted paragraph.");
        // Delete a run to create a deletion revision by Bob.
        if (doc.FirstSection.Body.Paragraphs.Count > 0)
        {
            Paragraph para = doc.FirstSection.Body.Paragraphs[0];
            if (para.Runs.Count > 0)
                para.Runs[0].Remove();
        }
        doc.StopTrackRevisions();

        // Save the sample document.
        doc.Save(filePath);
    }
}
