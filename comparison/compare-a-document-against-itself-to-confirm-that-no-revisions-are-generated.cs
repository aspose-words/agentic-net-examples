using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with deterministic content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph for self‑comparison.");

        // Compare the document with itself. No revisions should be generated.
        doc.Compare(doc, "SelfComparer", DateTime.Now);

        // Verify that the revisions collection is empty.
        if (doc.Revisions.Count != 0)
        {
            throw new InvalidOperationException(
                $"Expected zero revisions, but found {doc.Revisions.Count}.");
        }

        // Save the resulting document (it will be identical to the original).
        string outputPath = "self-compare.docx";
        doc.Save(outputPath);
    }
}
