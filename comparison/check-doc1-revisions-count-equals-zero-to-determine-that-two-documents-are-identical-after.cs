using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the first document with some content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Hello world.");

        // Create the second document with identical content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Hello world.");

        // Ensure both documents have no revisions before comparison.
        if (doc1.Revisions.Count != 0 || doc2.Revisions.Count != 0)
            throw new InvalidOperationException("Documents should not contain revisions before comparison.");

        // Compare the documents. Since they are identical, no revisions should be generated.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Verify that the comparison produced zero revisions.
        if (doc1.Revisions.Count != 0)
            throw new InvalidOperationException("Expected zero revisions for identical documents.");

        // Save the (unchanged) result document.
        doc1.Save("identical-comparison.docx");

        // Indicate success.
        Console.WriteLine("Documents are identical; no revisions were created.");
    }
}
