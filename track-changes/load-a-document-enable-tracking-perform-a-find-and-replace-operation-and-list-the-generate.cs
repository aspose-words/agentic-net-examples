using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        var builder = new DocumentBuilder();
        builder.Writeln("Hello World! This is a sample document for track changes demonstration.");

        // Save the sample document to a temporary file.
        string originalPath = "Original.docx";
        builder.Document.Save(originalPath);

        // Load the document from the file.
        Document doc = new Document(originalPath);

        // Enable track changes with a specific author and current date.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Perform a find-and-replace operation while tracking is enabled.
        // This will generate revisions for the replaced text.
        FindReplaceOptions options = new FindReplaceOptions();
        doc.Range.Replace("Hello", "Hi", options);

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // List all generated revisions.
        var revisions = doc.Revisions;
        if (revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were generated.");
        }

        Console.WriteLine($"Total revisions: {revisions.Count}");
        foreach (var rev in revisions)
        {
            string revType = rev.RevisionType.ToString();
            string author = rev.Author;
            DateTime date = rev.DateTime;
            string text = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;

            Console.WriteLine($"Type: {revType}, Author: {author}, Date: {date}, Text: \"{text}\"");
        }

        // Save the revised document.
        string revisedPath = "Revised.docx";
        doc.Save(revisedPath);
    }
}
