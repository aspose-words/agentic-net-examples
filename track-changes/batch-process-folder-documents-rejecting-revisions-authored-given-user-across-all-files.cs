using System;
using System.IO;
using Aspose.Words;

// Custom criteria that matches revisions authored by a specific user.
class AuthorRevisionCriteria : IRevisionCriteria
{
    private readonly string _authorName;

    public AuthorRevisionCriteria(string authorName)
    {
        _authorName = authorName;
    }

    // Returns true only for revisions whose Author property matches the specified name.
    public bool IsMatch(Revision revision)
    {
        return revision.Author == _authorName;
    }
}

class BatchRejectRevisions
{
    static void Main()
    {
        // Folder containing the source documents.
        string sourceFolder = @"C:\Docs\Input";

        // Folder where the processed documents will be saved.
        string targetFolder = @"C:\Docs\Output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(targetFolder);

        // Author whose revisions should be rejected.
        string authorToReject = "John Doe";

        // Process each .docx file in the source folder.
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the document (lifecycle rule: load).
            Document doc = new Document(filePath);

            // Reject all revisions that match the custom criteria (author name).
            doc.Revisions.Reject(new AuthorRevisionCriteria(authorToReject));

            // Save the modified document (lifecycle rule: save).
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(targetFolder, fileName);
            doc.Save(outputPath);
        }
    }
}
