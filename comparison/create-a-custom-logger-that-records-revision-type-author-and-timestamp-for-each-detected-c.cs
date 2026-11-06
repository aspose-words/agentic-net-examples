using System;
using System.IO;
using Aspose.Words;

public class RevisionLoggerExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the first paragraph.");
        builderOriginal.Writeln("This is the second paragraph.");
        builderOriginal.Writeln("This is the third paragraph.");

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(outputDir, "Original.docx");
        original.Save(originalPath);

        // Create the revised document with intentional changes:
        // - second paragraph modified,
        // - third paragraph omitted (deleted),
        // - a new fourth paragraph added.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the first paragraph."); // unchanged
        builderRevised.Writeln("This is the modified second paragraph."); // changed
        // third paragraph omitted to simulate deletion
        builderRevised.Writeln("This is an added fourth paragraph."); // new

        // Save the revised document (optional).
        string revisedPath = Path.Combine(outputDir, "Revised.docx");
        revised.Save(revisedPath);

        // Perform comparison, providing author name and comparison date.
        string comparisonAuthor = "RevisionLogger";
        DateTime comparisonDate = DateTime.Now;
        original.Compare(revised, comparisonAuthor, comparisonDate);

        // Save the compared document that now contains revisions.
        string comparedPath = Path.Combine(outputDir, "Compared.docx");
        original.Save(comparedPath);

        // Prepare the log file.
        string logPath = Path.Combine(outputDir, "RevisionLog.txt");
        using StreamWriter logWriter = new StreamWriter(logPath, false);
        logWriter.WriteLine("RevisionType,Author,Timestamp");

        // Iterate through revisions and log details.
        foreach (Revision revision in original.Revisions)
        {
            string type = revision.RevisionType.ToString();
            string author = revision.Author ?? "Unknown";
            DateTime timestamp = revision.DateTime;

            string line = $"{type},{author},{timestamp:O}";
            logWriter.WriteLine(line);
        }

        // Ensure the log is flushed.
        logWriter.Flush();

        // Output summary to the console.
        Console.WriteLine($"Comparison completed. Revisions logged to: {logPath}");
        Console.WriteLine($"Total revisions detected: {original.Revisions.Count}");
    }
}
