using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This is the original paragraph.");

        // Create the revised document with differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello world!"); // punctuation change
        builderRevised.Writeln("This is the revised paragraph."); // text change

        // Compare the documents. Revisions will be added to the original document.
        string comparisonAuthor = "Comparer";
        DateTime comparisonDate = DateTime.Now;
        original.Compare(revised, comparisonAuthor, comparisonDate);

        // Ensure that revisions were detected.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after comparison.");
        }

        // Prepare a log file to record revision details.
        string logFilePath = Path.Combine(Directory.GetCurrentDirectory(), "revision_log.txt");
        using (StreamWriter logger = new StreamWriter(logFilePath, false))
        {
            logger.WriteLine("RevisionType\tAuthor\tTimestamp");
            foreach (Revision rev in original.Revisions)
            {
                // Log revision type, author, and timestamp in ISO 8601 format.
                logger.WriteLine($"{rev.RevisionType}\t{rev.Author}\t{rev.DateTime:O}");
            }
        }

        // Save the compared document for reference.
        string comparedDocPath = Path.Combine(Directory.GetCurrentDirectory(), "compared.docx");
        original.Save(comparedDocPath);
    }
}
