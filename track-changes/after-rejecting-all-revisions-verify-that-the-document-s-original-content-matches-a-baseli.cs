using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        string baselinePath = "baseline.docx";
        string resultPath = "result.docx";

        // Step 1: Create a baseline document with original content.
        Document baselineDoc = new Document();
        DocumentBuilder baselineBuilder = new DocumentBuilder(baselineDoc);
        baselineBuilder.Writeln("Hello World!");
        baselineDoc.Save(baselinePath);

        // Step 2: Load the baseline and start tracking revisions.
        Document doc = new Document(baselinePath);
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Step 3: Make a change that generates a revision (insert text).
        DocumentBuilder revBuilder = new DocumentBuilder(doc);
        revBuilder.MoveToDocumentEnd();
        revBuilder.Writeln("This is an inserted line.");

        // Step 4: Stop tracking revisions.
        doc.StopTrackRevisions();

        // Step 5: Reject all revisions, reverting to the original content.
        // Use the RevisionCollection API (AcceptAll/RejectAll) as required by Aspose.Words.
        doc.Revisions.RejectAll();

        // Save the resulting document.
        doc.Save(resultPath);

        // Step 6: Verify that the resulting document matches the baseline.
        Document resultDoc = new Document(resultPath);
        string baselineText = baselineDoc.GetText();
        string resultText = resultDoc.GetText();

        if (baselineText == resultText)
        {
            Console.WriteLine("Verification succeeded: the document matches the baseline.");
        }
        else
        {
            throw new InvalidOperationException("Verification failed: the document does not match the baseline.");
        }

        // Optional cleanup.
        // File.Delete(baselinePath);
        // File.Delete(resultPath);
    }
}
