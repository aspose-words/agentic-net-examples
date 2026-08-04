using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This is a sample document.");
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        doc.Save(originalPath);

        // Load the document.
        Document loadedDoc = new Document(originalPath);

        // Enable track changes.
        loadedDoc.StartTrackRevisions("Alice", DateTime.Now);

        // Perform a find-and-replace operation that will be recorded as a revision.
        loadedDoc.Range.Replace("Hello", "Hi", new FindReplaceOptions());

        // Stop tracking revisions.
        loadedDoc.StopTrackRevisions();

        // List generated revisions.
        Console.WriteLine($"Total revisions: {loadedDoc.Revisions.Count}");
        foreach (Revision rev in loadedDoc.Revisions)
        {
            string text = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}, Text: \"{text}\"");
        }

        // Save the modified document.
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        loadedDoc.Save(revisedPath);
    }
}
