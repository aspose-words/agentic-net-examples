using System;
using System.IO;
using Aspose.Words;

public class CustomRevisionLogger
{
    public static void Main()
    {
        // Prepare a folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonArtifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");
        builderOriginal.Writeln("This line will stay unchanged.");

        // Save the original (optional, for inspection).
        string originalPath = Path.Combine(artifactsDir, "original.docx");
        original.Save(originalPath);

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words!"); // changed text
        builderRevised.Writeln("This line will stay unchanged."); // same line
        builderRevised.Writeln("An extra line was added."); // new line

        // Save the revised (optional, for inspection).
        string revisedPath = Path.Combine(artifactsDir, "revised.docx");
        revised.Save(revisedPath);

        // Perform comparison: original will receive revisions.
        string author = "CustomLogger";
        DateTime compareTime = DateTime.Now;
        original.Compare(revised, author, compareTime);

        // Verify that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were detected after comparison.");

        // Prepare a log for revisions.
        string logPath = Path.Combine(artifactsDir, "revision_log.txt");
        using (StreamWriter logWriter = new StreamWriter(logPath, false))
        {
            logWriter.WriteLine("RevisionType\tAuthor\tTimestamp");
            foreach (Revision rev in original.Revisions)
            {
                // Ensure properties are not null (they are guaranteed by Aspose.Words).
                string line = $"{rev.RevisionType}\t{rev.Author}\t{rev.DateTime:u}";
                logWriter.WriteLine(line);
                Console.WriteLine(line);
            }
        }

        // Save the compared document showing revisions.
        string comparedPath = Path.Combine(artifactsDir, "compared.docx");
        original.Save(comparedPath);
    }
}
