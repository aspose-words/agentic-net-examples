using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        string inputPath = "RevisionsInput.docx";
        string outputPath = "RevisionsCleaned.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains tracked changes.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Normal text – not a revision.
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions, add some text, then stop tracking.
        sampleDoc.StartTrackRevisions("SampleAuthor");
        builder.Writeln("This line was added while tracking changes.");
        sampleDoc.StopTrackRevisions();

        // Save the document that now contains revisions.
        sampleDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document with revisions, accept all of them, and save.
        // -----------------------------------------------------------------
        Document docWithRevisions = new Document(inputPath);

        // Ensure that revisions are present; otherwise the task would be meaningless.
        if (!docWithRevisions.HasRevisions)
            throw new InvalidOperationException("The loaded document does not contain any revisions.");

        // Accept every revision in the document.
        docWithRevisions.AcceptAllRevisions();

        // Save the cleaned document to a new file.
        docWithRevisions.Save(outputPath);
    }
}
