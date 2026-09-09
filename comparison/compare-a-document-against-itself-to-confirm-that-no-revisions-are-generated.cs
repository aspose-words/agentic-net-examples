using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document with deterministic content.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is a sample paragraph for comparison.");

        // Clone the document to obtain a separate instance with identical content.
        Document clone = (Document)original.Clone(true);

        // Perform the comparison. No revisions should be generated because the contents are identical.
        original.Compare(clone, "Tester", DateTime.Now);

        // Verify that the comparison produced zero revisions.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException(
                $"Expected zero revisions, but found {original.Revisions.Count}.");
        }

        // Save the (unchanged) document to demonstrate that the operation completed successfully.
        original.Save("ComparisonResult.docx");
    }
}
