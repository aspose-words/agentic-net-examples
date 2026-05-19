using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used for self‑comparison.");

        // Compare the document with itself. No revisions should be generated.
        doc.Compare(doc, "SelfComparer", DateTime.Now);

        // Verify that the comparison produced zero revisions.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException(
                $"Expected zero revisions after self‑comparison, but found {doc.Revisions.Count}.");

        // Save the (unchanged) document to the current directory.
        string outputPath = "SelfCompared.docx";
        doc.Save(outputPath);

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Self‑comparison completed. Revisions count: {doc.Revisions.Count}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
