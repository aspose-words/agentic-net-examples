using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string baselinePath = Path.Combine(Directory.GetCurrentDirectory(), "Baseline.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a baseline document with original content.
        // -----------------------------------------------------------------
        Document baselineDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(baselineDoc);
        builder.Writeln("This is the original content.");
        baselineDoc.Save(baselinePath);

        // -----------------------------------------------------------------
        // Step 2: Load the baseline document and introduce revisions.
        // -----------------------------------------------------------------
        Document doc = new Document(baselinePath);
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Insert a new paragraph (creates an insertion revision).
        builder = new DocumentBuilder(doc);
        builder.Writeln("This line was added as a revision.");

        // Delete the original paragraph's first run (creates a deletion revision).
        if (doc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        doc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // Step 3: Reject all revisions, reverting the document to its original state.
        // -----------------------------------------------------------------
        doc.Revisions.RejectAll();

        // -----------------------------------------------------------------
        // Step 4: Verify that the document's content matches the baseline.
        // -----------------------------------------------------------------
        string finalText = doc.GetText().Trim();
        Document baselineCheck = new Document(baselinePath);
        string baselineText = baselineCheck.GetText().Trim();

        if (!finalText.Equals(baselineText, StringComparison.Ordinal))
            throw new Exception("The document content does not match the baseline after rejecting revisions.");

        // -----------------------------------------------------------------
        // Step 5: Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}
