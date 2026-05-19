using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the first document with deterministic content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is a sample paragraph used for comparison.");

        // Create the second document with exactly the same content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("This is a sample paragraph used for comparison.");

        // Perform the comparison. No revisions should be generated because the documents are identical.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Verify that the revisions collection is empty.
        if (doc1.Revisions.Count != 0)
            throw new InvalidOperationException("Documents are not identical – revisions were detected.");

        // Save the (unchanged) document to demonstrate that the operation completed successfully.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "IdenticalComparison.docx");
        doc1.Save(outputPath);
    }
}
