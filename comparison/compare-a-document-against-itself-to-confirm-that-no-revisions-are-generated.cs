using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document and add some deterministic content.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello world!");
        builder.Writeln("This is a test document.");

        // Ensure the original document has no revisions before comparison.
        if (docOriginal.HasRevisions)
            docOriginal.Revisions.AcceptAll();

        // Create an identical copy of the original document.
        Document docCopy = (Document)docOriginal.Clone(true);

        // Compare the original document with its identical copy.
        // No revisions should be generated because the contents are the same.
        docOriginal.Compare(docCopy, "Comparer", DateTime.Now);

        // Verify that the revision count is zero.
        int revisionCount = docOriginal.Revisions.Count;

        // Save the resulting document for inspection (optional artifact).
        docOriginal.Save("ZeroDifferenceResult.docx");

        // Output the verification result.
        Console.WriteLine($"Revisions after self‑comparison: {revisionCount}");
    }
}
