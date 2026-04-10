using System;
using System.IO;
using Aspose.Words;

public class CompareIdenticalDocuments
{
    public static void Main()
    {
        // Define a folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the first document and add identical content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is a sample paragraph.");
        builder1.Writeln("Another line of text.");

        // Create the second document with the same content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("This is a sample paragraph.");
        builder2.Writeln("Another line of text.");

        // Ensure both documents have no revisions before comparison.
        if (doc1.Revisions.Count != 0 || doc2.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");
        }

        // Compare the documents. Since they are identical, no revisions should be created.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Verify that the revisions count is zero, indicating identical documents.
        bool areIdentical = doc1.Revisions.Count == 0;
        Console.WriteLine($"Documents are identical: {areIdentical}");

        // Save the (unchanged) document to demonstrate that an artifact is produced.
        string resultPath = Path.Combine(outputDir, "ComparisonResult.docx");
        doc1.Save(resultPath);
    }
}
