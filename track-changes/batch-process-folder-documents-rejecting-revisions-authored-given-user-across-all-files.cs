using System;
using System.IO;
using Aspose.Words;

class AuthorRevisionCriteria : IRevisionCriteria
{
    private readonly string _authorName;
    public AuthorRevisionCriteria(string authorName) => _authorName = authorName;
    public bool IsMatch(Revision revision) => revision.Author == _authorName;
}

class BatchRejectRevisions
{
    static void Main()
    {
        string baseDir = AppContext.BaseDirectory;
        string sourceFolder = Path.Combine(baseDir, "Input");
        string targetFolder = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(targetFolder);

        string authorToReject = "John Doe";

        var files = Directory.GetFiles(sourceFolder, "*.docx");
        if (files.Length == 0)
        {
            Console.WriteLine($"No .docx files found in '{sourceFolder}'. Place files there and rerun.");
            return;
        }

        foreach (var filePath in files)
        {
            Document doc = new Document(filePath);
            doc.Revisions.Reject(new AuthorRevisionCriteria(authorToReject));

            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(targetFolder, fileName);
            doc.Save(outputPath);
            Console.WriteLine($"Processed '{fileName}' -> '{outputPath}'.");
        }
    }
}
