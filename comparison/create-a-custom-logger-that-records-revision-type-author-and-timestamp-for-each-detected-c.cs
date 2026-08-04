using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class RevisionLogger
{
    private readonly List<string> _entries = new();

    public void Log(Revision revision)
    {
        // Record revision type, author and timestamp in a readable format.
        string line = $"{revision.RevisionType}\t{revision.Author}\t{revision.DateTime:u}";
        _entries.Add(line);
    }

    public void Save(string filePath)
    {
        // Write all logged entries to a text file.
        File.WriteAllLines(filePath, _entries);
    }
}

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("The quick brown fox jumps over the lazy dog.");
        builderOriginal.Writeln("This line will stay unchanged.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("The quick brown fox jumps over the energetic cat."); // changed word
        builderRevised.Writeln("This line will stay unchanged."); // same line
        builderRevised.Writeln("An additional line is added."); // new line

        // Perform comparison. Use a distinct author name.
        string author = "Alice";
        DateTime compareTime = DateTime.Now;
        original.Compare(revised, author, compareTime);

        // Ensure that revisions were detected.
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
        string outputDoc = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        string logFile = Path.Combine(Directory.GetCurrentDirectory(), "revision_log.txt");

        original.Save(outputDoc);
        logger.Save(logFile);
    }
}
