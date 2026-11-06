using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the first document with deterministic content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Identical line 1.");
        builder1.Writeln("Identical line 2.");

        // Create the second document with the same content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Identical line 1.");
        builder2.Writeln("Identical line 2.");

        // Compare the two documents. No differences should be detected.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Verify that the comparison produced zero revisions.
        if (doc1.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Expected zero revisions for identical documents.");
        }

        // Save the (unchanged) document to demonstrate successful execution.
        doc1.Save("identical-comparison.docx");
    }
}
