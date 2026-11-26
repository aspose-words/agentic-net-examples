using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class RevisionLogger
{
    private readonly List<string> _entries = new();

    // Logs the revision type, author and timestamp.
    public void Log(Revision revision)
    {
        string entry = $"{revision.RevisionType},{revision.Author},{revision.DateTime:O}";
        _entries.Add(entry);
    }

    // Saves the collected log entries to a CSV file.
    public void Save(string filePath)
    {
        string? directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using StreamWriter writer = new(filePath, false);
        writer.WriteLine("RevisionType,Author,Timestamp");
        foreach (string entry in _entries)
        {
            writer.WriteLine(entry);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // ---------- Create original document ----------
        Document original = new();
        DocumentBuilder builderOriginal = new(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("Another line that will stay unchanged.");

        // ---------- Create revised document with changes ----------
        Document revised = new();
        DocumentBuilder builderRevised = new(revised);
        builderRevised.Writeln("This is the modified paragraph."); // Changed text.
        builderRevised.Writeln("Another line that will stay unchanged."); // Same text.
        builderRevised.Writeln("An added line."); // New line.

        // ---------- Perform comparison ----------
        string author = "ComparisonAuthor";
        DateTime compareDate = DateTime.Now;
        original.Compare(revised, author, compareDate);

        // Verify that revisions were generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // ---------- Log revisions ----------
        RevisionLogger logger = new();
        foreach (Revision rev in original.Revisions)
        {
            logger.Log(rev);
        }

        // ---------- Save compared document ----------
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(comparedPath);

        // ---------- Save revision log ----------
        string logPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionLog.csv");
        logger.Save(logPath);
    }
}
