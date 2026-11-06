using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string baselinePath = "Baseline.docx";
        const string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // Create baseline document with original content
        // -----------------------------------------------------------------
        Document baselineDoc = new Document();
        DocumentBuilder baselineBuilder = new DocumentBuilder(baselineDoc);
        baselineBuilder.Writeln("Hello world.");
        baselineDoc.Save(baselinePath);

        // -----------------------------------------------------------------
        // Load baseline and start tracking revisions
        // -----------------------------------------------------------------
        Document doc = new Document(baselinePath);
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Make some changes that will be recorded as revisions
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert a new paragraph (revision: insertion)
        builder.Writeln("This is a new paragraph added while tracking.");

        // Replace text (revision: deletion + insertion)
        // Use FindReplaceOptions with a regular expression
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        doc.Range.Replace(new Regex("Hello"), "Hi", replaceOptions);

        // Stop tracking revisions
        doc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // Reject all revisions, restoring the document to its original state
        // -----------------------------------------------------------------
        // The RevisionCollection provides RejectAll() to reject every revision.
        doc.Revisions.RejectAll();

        // Save the document after rejection
        doc.Save(resultPath);

        // -----------------------------------------------------------------
        // Verify that the content matches the baseline document
        // -----------------------------------------------------------------
        Document afterReject = new Document(resultPath);
        Document originalBaseline = new Document(baselinePath);

        string afterText = afterReject.GetText();
        string baselineText = originalBaseline.GetText();

        if (afterText != baselineText)
        {
            throw new InvalidOperationException("Document content does not match the baseline after rejecting revisions.");
        }

        // If we reach this point, verification succeeded
        Console.WriteLine("Revisions rejected successfully; document matches the baseline.");
    }
}
