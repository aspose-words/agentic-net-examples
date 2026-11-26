using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for files used in the example
        string baselinePath = "Baseline.docx";
        string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a baseline document with original content.
        // -----------------------------------------------------------------
        Document baselineDoc = new Document();
        DocumentBuilder baselineBuilder = new DocumentBuilder(baselineDoc);
        baselineBuilder.Writeln("Hello World!");
        baselineDoc.Save(baselinePath);

        // -----------------------------------------------------------------
        // Step 2: Load the baseline and make changes while tracking revisions.
        // -----------------------------------------------------------------
        Document doc = new Document(baselinePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking of revisions.
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Insert a new paragraph (this will be recorded as an insertion revision).
        builder.Writeln("This is a newly added line.");

        // Modify existing text: prepend "Hi " to the first paragraph.
        builder.MoveToDocumentStart();
        builder.Write("Hi ");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // Step 3: Reject all revisions, reverting the document to its original state.
        // -----------------------------------------------------------------
        // If the RejectAllRevisions method is unavailable, reject each revision manually.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            doc.Revisions[i].Reject();
        }

        // Save the document after rejecting revisions.
        doc.Save(resultPath);

        // -----------------------------------------------------------------
        // Step 4: Verify that the resulting document matches the baseline.
        // -----------------------------------------------------------------
        string baselineText = new Document(baselinePath).GetText();
        string resultText = new Document(resultPath).GetText();

        if (!string.Equals(baselineText, resultText, StringComparison.Ordinal))
        {
            throw new Exception("The document content does not match the baseline after rejecting revisions.");
        }

        // If no exception was thrown, verification succeeded.
        Console.WriteLine("Verification succeeded: the document matches the baseline after rejecting revisions.");
    }
}
