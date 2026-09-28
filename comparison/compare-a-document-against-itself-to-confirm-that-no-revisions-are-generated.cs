using System;
using Aspose.Words;

public class SelfComparisonExample
{
    public static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph for self‑comparison.");

        // Compare the document with itself. No revisions should be generated.
        doc.Compare(doc, "SelfCompareAuthor", DateTime.Now);

        // Verify that the revisions collection is empty.
        if (doc.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Expected zero revisions after comparing the document with itself.");
        }

        // Save the resulting document (it will be identical to the original).
        doc.Save("SelfComparisonResult.docx");
    }
}
