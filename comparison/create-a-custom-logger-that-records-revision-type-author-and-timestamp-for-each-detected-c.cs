using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

public class RevisionLogger
{
    private readonly StringBuilder _logBuilder = new StringBuilder();

    public void Log(Revision revision)
    {
        // Record revision type, author and timestamp in ISO 8601 format.
        _logBuilder.AppendLine($"{revision.RevisionType}\t{revision.Author}\t{revision.DateTime:O}");
    }

    public void Save(string filePath)
    {
        File.WriteAllText(filePath, _logBuilder.ToString());
    }
}

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");
        builderOriginal.Writeln("This line will stay unchanged.");

        // Create the revised document with some modifications.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words!"); // Modified text.
        builderRevised.Writeln("This line will stay unchanged.");
        builderRevised.Writeln("An extra line added."); // Insertion.

        // Perform comparison. Author and timestamp are required.
        string author = "Comparer";
        DateTime compareTime = DateTime.Now;
        original.Compare(revised, author, compareTime);

        // Verify that revisions were detected.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after comparison.");
        }

        // Log each revision's details.
        RevisionLogger logger = new RevisionLogger();
        foreach (Revision rev in original.Revisions)
        {
            logger.Log(rev);
        }

        // Save the compared document and the revision log.
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "compared.docx");
        original.Save(outputDocPath);

        string logPath = Path.Combine(Directory.GetCurrentDirectory(), "revision_log.txt");
        logger.Save(logPath);
    }
}
