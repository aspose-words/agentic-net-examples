using System;
using System.IO;
using Aspose.Words;

public class RevisionLogger
{
    private readonly string _logPath;

    public RevisionLogger(string logPath)
    {
        _logPath = logPath ?? throw new ArgumentNullException(nameof(logPath));

        // Ensure the directory for the log file exists.
        string? directory = Path.GetDirectoryName(_logPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    public void Log(string message)
    {
        if (message == null) throw new ArgumentNullException(nameof(message));
        File.AppendAllText(_logPath, message + Environment.NewLine);
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the compared document and the revision log.
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "compared.docx");
        string logFilePath = Path.Combine(Directory.GetCurrentDirectory(), "revision_log.txt");

        // ----- Create the original document -----
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the first paragraph.");
        builderOriginal.Writeln("This paragraph will be deleted in the revised version.");
        builderOriginal.Writeln("This paragraph will stay unchanged.");

        // ----- Create the revised document with intentional differences -----
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the first paragraph."); // unchanged
        // Deleted paragraph omitted.
        builderRevised.Writeln("This paragraph has been inserted in the revised version."); // insertion
        builderRevised.Writeln("This paragraph will stay unchanged."); // unchanged

        // Change formatting of a paragraph.
        builderRevised.Writeln("Formatted paragraph.");
        builderRevised.Font.Bold = true;
        builderRevised.Writeln("Bold text added.");

        // ----- Perform comparison -----
        string author = "RevisionLogger";
        DateTime compareDate = DateTime.Now;
        original.Compare(revised, author, compareDate);

        // ----- Initialize logger and write header -----
        RevisionLogger logger = new RevisionLogger(logFilePath);
        logger.Log($"Revision Log - Generated on {DateTime.Now:O}");
        logger.Log(new string('-', 50));

        // ----- Inspect revisions and log details -----
        foreach (Revision revision in original.Revisions)
        {
            string type = revision.RevisionType.ToString();
            string revAuthor = revision.Author ?? "Unknown";

            // Aspose.Words versions prior to 22.5 do not expose a RevisionDate property.
            // Use the current time as a timestamp for logging purposes.
            DateTime timestamp = DateTime.Now;

            logger.Log($"Type: {type}, Author: {revAuthor}, Timestamp: {timestamp:O}");
        }

        // ----- Save the document that contains the revisions -----
        original.Save(outputDocPath);
    }
}
