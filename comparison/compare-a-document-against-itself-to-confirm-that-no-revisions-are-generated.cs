using System;
using Aspose.Words;

public class SelfComparisonExample
{
    public static void Main()
    {
        // Create a blank document and add some content.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is a sample paragraph for self‑comparison.");

        // Clone the original document to obtain a separate instance with identical content.
        Document clone = (Document)original.Clone(true);

        // Compare the original document with its clone.
        // Since the contents are identical, no revisions should be generated.
        original.Compare(clone, "SelfComparer", DateTime.Now);

        // Verify that the comparison produced zero revisions.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were generated when comparing identical documents.");

        // Save the (unchanged) document to demonstrate that the operation completed successfully.
        string outputPath = "self_compare.docx";
        original.Save(outputPath);
    }
}
