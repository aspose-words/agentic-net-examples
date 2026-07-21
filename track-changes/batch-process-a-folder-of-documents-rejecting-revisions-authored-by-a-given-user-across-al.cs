using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Criteria to match revisions authored by a specific user.
    private class RevisionAuthorCriteria : IRevisionCriteria
    {
        private readonly string _author;

        public RevisionAuthorCriteria(string author)
        {
            _author = author;
        }

        public bool IsMatch(Revision revision)
        {
            return revision.Author == _author;
        }
    }

    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Author whose revisions will be rejected.
        string targetAuthor = "Bob";

        // Create sample documents containing revisions from two authors.
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"));

        // Process each document: reject revisions authored by the target author.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Reject matching revisions; the method returns the number of rejected revisions.
            doc.Revisions.Reject(new RevisionAuthorCriteria(targetAuthor));

            // Save the cleaned document.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }

    // Generates a document with revisions from "Alice" and "Bob".
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normal text (no revision).
        builder.Writeln("Initial content. ");

        // Revisions by Alice.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Added by Alice. ");
        doc.StopTrackRevisions();

        // Revisions by Bob.
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Added by Bob. ");

        // Delete a run to create a deletion revision by Bob.
        if (doc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        doc.StopTrackRevisions();

        doc.Save(filePath);
    }
}
