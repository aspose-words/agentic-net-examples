using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string samplePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // Create a simple document and save it – this acts as the source file.
        // -----------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder creator = new DocumentBuilder(createDoc);
        creator.Writeln("Hello world! This is a sample document for tracking changes.");
        createDoc.Save(samplePath);

        // -------------------------------------------------
        // Load the document we just created.
        // -------------------------------------------------
        Document doc = new Document(samplePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Enable track revisions with a specific author and date.
        // -------------------------------------------------
        string author = "AsposeUser";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // -------------------------------------------------
        // Perform a find-and-replace operation while tracking is enabled.
        // This will generate a revision.
        // -------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions();
        doc.Range.Replace("world", "Aspose", options);

        // -------------------------------------------------
        // Stop tracking further changes.
        // -------------------------------------------------
        doc.StopTrackRevisions();

        // -------------------------------------------------
        // List all generated revisions.
        // -------------------------------------------------
        Console.WriteLine($"Total revisions: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            string revText = rev.ParentNode?.GetText().Trim() ?? string.Empty;
            Console.WriteLine($"Revision Type: {rev.RevisionType}");
            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime}");
            Console.WriteLine($"Text: \"{revText}\"");
            Console.WriteLine(new string('-', 40));
        }

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}
