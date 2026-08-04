using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Criteria that matches revisions authored by a specific user.
    private class RevisionAuthorCriteria : IRevisionCriteria
    {
        private readonly string _author;
        public RevisionAuthorCriteria(string author) => _author = author;
        public bool IsMatch(Revision revision) => revision.Author == _author;
    }

    public static void Main()
    {
        // Folder paths (relative to the executable directory).
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");
        string targetAuthor = "John Doe";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If the input folder is empty, create sample documents with revisions.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            CreateSampleDocument(Path.Combine(inputFolder, "DocWithTargetAuthor.docx"), targetAuthor, "Text added by target author.");
            CreateSampleDocument(Path.Combine(inputFolder, "DocWithOtherAuthor.docx"), "Jane Smith", "Text added by other author.");
        }

        // Process each .docx file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Reject all revisions authored by the target user.
            doc.Revisions.Reject(new RevisionAuthorCriteria(targetAuthor));

            // Save the processed document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }

    // Helper method to create a sample document with a single revision by the specified author.
    private static void CreateSampleDocument(string filePath, string author, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions with the given author.
        doc.StartTrackRevisions(author, DateTime.Now);
        builder.Writeln(text);
        doc.StopTrackRevisions();

        // Save the sample document.
        doc.Save(filePath);
    }
}
