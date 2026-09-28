using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document. Hello World!");

        // Save the sample document locally.
        string samplePath = "Sample.docx";
        doc.Save(samplePath);

        // Load the document from the file.
        Document loadedDoc = new Document(samplePath);

        // Enable track changes with an author name and current date.
        loadedDoc.StartTrackRevisions("John Doe", DateTime.Now);

        // Perform a find-and-replace operation while tracking is enabled.
        loadedDoc.Range.Replace("World", "Aspose", new FindReplaceOptions());

        // Stop tracking changes.
        loadedDoc.StopTrackRevisions();

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        loadedDoc.Save(modifiedPath);

        // List all generated revisions.
        foreach (Revision rev in loadedDoc.Revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}");
            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime}");
            Console.WriteLine($"Text: {rev.ParentNode?.GetText().Trim()}");
            Console.WriteLine("---");
        }
    }
}
