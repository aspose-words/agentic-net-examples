using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It contains some sample text.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the edited paragraph."); // changed text
        builderRevised.Writeln("It contains some sample text."); // unchanged
        builderRevised.Writeln("An additional line was added."); // new line

        // Perform comparison. The revisions will be stored in the original document.
        string author = "CustomLogger";
        DateTime compareTime = DateTime.Now;
        original.Compare(revised, author, compareTime);

        // Prepare a logger to capture revision details.
        List<string> logLines = new List<string>
        {
            $"Comparison performed by '{author}' at {compareTime:u}",
            $"Total revisions detected: {original.Revisions.Count}"
        };

        // Iterate through each revision and record its type, author, and timestamp.
        foreach (Revision rev in original.Revisions)
        {
            string line = $"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Timestamp: {rev.DateTime:u}";
            logLines.Add(line);
        }

        // Save the compared document with revisions.
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "compared.docx");
        original.Save(outputDocPath);

        // Save the revision log to a text file.
        string logPath = Path.Combine(Directory.GetCurrentDirectory(), "revision_log.txt");
        File.WriteAllLines(logPath, logLines);
    }
}
