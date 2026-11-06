using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with deterministic content.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Sample text for self-comparison.");

        // Clone the document to obtain an identical copy.
        Document clone = (Document)original.Clone(true);

        // Compare the original document against its identical clone.
        original.Compare(clone, "SelfCompareAuthor", DateTime.Now);

        // Verify that no revisions were generated.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException($"Expected zero revisions, but found {original.Revisions.Count}.");
        }

        // Save the result document (it will be identical to the original).
        original.Save("self-compare-result.docx");
    }
}
