using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the first document with deterministic content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Sample text for comparison.");

        // Create the second document with identical content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Sample text for comparison.");

        // Perform comparison. Since the documents are identical, no revisions should be generated.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Verify that the revisions collection is empty.
        if (doc1.Revisions.Count != 0)
        {
            throw new InvalidOperationException(
                $"Expected zero revisions, but found {doc1.Revisions.Count}.");
        }

        // Save the (unchanged) document to demonstrate that the process completed.
        doc1.Save("identical-comparison.docx");
    }
}
