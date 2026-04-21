using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace TrackChangesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with some text.
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("The quick brown fox jumps over the lazy dog.");
            builder.Writeln("Another line with the word fox.");
            // Save the sample document to disk.
            const string samplePath = "sample.docx";
            sampleDoc.Save(samplePath);

            // Load the document from disk.
            Document doc = new Document(samplePath);

            // Enable tracking of revisions.
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            // Perform a find-and-replace operation while tracking is enabled.
            FindReplaceOptions replaceOptions = new FindReplaceOptions();
            // Replace the word "fox" with "cat". This will generate deletion and insertion revisions.
            doc.Range.Replace("fox", "cat", replaceOptions);

            // Stop tracking further changes.
            doc.StopTrackRevisions();

            // Save the revised document.
            const string revisedPath = "revised.docx";
            doc.Save(revisedPath);

            // List all generated revisions.
            Console.WriteLine("Revisions generated in the document:");
            foreach (Revision rev in doc.Revisions)
            {
                // Output revision details: type, author, date, and the affected text.
                string text = rev.ParentNode?.GetText()?.Trim() ?? "<no text>";
                Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}, Text: \"{text}\"");
            }
        }
    }
}
