using System;
using System.IO;
using Aspose.Words;

public class TrackChangesDemo
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Step 1: Create a sample document and save it.
        string samplePath = Path.Combine(outputDir, "Sample.docx");
        Document sampleDoc = new Document();
        DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
        sampleBuilder.Writeln("This is the original paragraph.");
        sampleDoc.Save(samplePath);

        // Step 2: Load the document.
        Document doc = new Document(samplePath);

        // Step 3: Enable revision tracking.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Step 4: Insert new text (creates an insertion revision).
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph was inserted while tracking changes.");

        // Step 5: Apply a formatting change (won't be recorded as a revision by Aspose.Words).
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.Runs[0].Font.Bold = true;

        // Step 6: Delete the original paragraph (creates a deletion revision).
        firstParagraph.Remove();

        // Step 7: Stop tracking revisions.
        doc.StopTrackRevisions();

        // Step 8: Save the revised document.
        string revisedPath = Path.Combine(outputDir, "Revised.docx");
        doc.Save(revisedPath);

        // Step 9: List all revisions with their details.
        Console.WriteLine("Revisions found in the document:");
        foreach (Revision rev in doc.Revisions)
        {
            string text = rev.ParentNode != null ? rev.ParentNode.GetText().Trim() : "<no node>";
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}, Text: \"{text}\"");
        }
    }
}
