using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        string samplePath = "Sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This is a sample document.");
        doc.Save(samplePath);

        // Load the document.
        Document loadedDoc = new Document(samplePath);

        // Enable track changes.
        loadedDoc.StartTrackRevisions("John Doe", DateTime.Now);

        // Perform a find-and-replace operation.
        FindReplaceOptions options = new FindReplaceOptions();
        loadedDoc.Range.Replace("sample", "test", options);

        // Stop tracking revisions.
        loadedDoc.StopTrackRevisions();

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        loadedDoc.Save(modifiedPath);

        // List the generated revisions.
        RevisionCollection revisions = loadedDoc.Revisions;
        Console.WriteLine($"Total revisions: {revisions.Count}");
        foreach (Revision rev in revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }
    }
}
