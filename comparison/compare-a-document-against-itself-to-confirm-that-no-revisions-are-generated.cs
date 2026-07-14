using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class SelfComparisonExample
{
    public static void Main()
    {
        // Create a simple document with deterministic content.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is a sample paragraph used for self‑comparison.");

        // Compare the document with itself. No revisions should be generated.
        original.Compare(original, "SelfComparer", DateTime.Now);

        // Verify that the revisions collection is empty.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException(
                $"Expected zero revisions, but found {original.Revisions.Count}.");

        // Save the (unchanged) document to demonstrate that the operation completed.
        string outputPath = "SelfComparisonResult.docx";
        original.Save(outputPath);

        // Optional: write a brief confirmation to the console.
        Console.WriteLine($"Self‑comparison completed. Revisions count: {original.Revisions.Count}");
        Console.WriteLine($"Result saved to: {outputPath}");
    }
}
