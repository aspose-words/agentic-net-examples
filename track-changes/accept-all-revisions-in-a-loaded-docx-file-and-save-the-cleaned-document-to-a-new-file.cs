using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths for the sample input and the cleaned output.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "SampleWithRevisions.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CleanedDocument.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample document that contains tracked changes.
        // -----------------------------------------------------------------
        Document docWithRevisions = new Document();
        DocumentBuilder builder = new DocumentBuilder(docWithRevisions);

        // Normal text – not a revision.
        builder.Writeln("This paragraph is unchanged.");

        // Start tracking revisions.
        docWithRevisions.StartTrackRevisions("Author", DateTime.Now);

        // Insert text that will be recorded as an insertion revision.
        builder.Writeln("This paragraph was added while tracking changes.");

        // Delete a run to create a deletion revision.
        // For simplicity, remove the first run of the first paragraph.
        if (docWithRevisions.FirstSection.Body.FirstParagraph.Runs.Count > 0)
        {
            docWithRevisions.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        }

        // Stop tracking revisions.
        docWithRevisions.StopTrackRevisions();

        // Save the document that contains revisions.
        docWithRevisions.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document, accept all revisions, and save cleaned version.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Ensure the document actually has revisions before accepting.
        if (loadedDoc.HasRevisions)
        {
            loadedDoc.AcceptAllRevisions();
        }

        // Save the cleaned document without any tracked changes.
        loadedDoc.Save(outputPath);
    }
}
