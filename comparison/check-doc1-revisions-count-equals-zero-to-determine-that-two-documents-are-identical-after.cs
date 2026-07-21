using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the first document with deterministic content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is a sample paragraph for comparison.");
        builder1.Writeln("Another line with the same text.");

        // Create the second document with identical content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("This is a sample paragraph for comparison.");
        builder2.Writeln("Another line with the same text.");

        // Perform the comparison. No revisions should be generated because the documents are identical.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Verify that the revisions count is zero, indicating identical documents.
        if (doc1.Revisions.Count != 0)
            throw new InvalidOperationException("Documents differ: revisions were detected.");

        // Save the (unchanged) document to demonstrate that an output artifact is produced.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "IdenticalComparisonResult.docx");
        doc1.Save(outputPath);
    }
}
