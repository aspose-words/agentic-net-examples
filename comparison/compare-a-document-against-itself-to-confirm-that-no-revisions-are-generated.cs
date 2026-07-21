using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for self‑comparison.");

        // The document must not contain revisions before calling Compare.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException("Document should have no revisions before comparison.");

        // Compare the document with itself. No revisions should be produced.
        doc.Compare(doc, "Author", DateTime.Now);

        // Verify that the comparison did not generate any revisions.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were generated when comparing identical documents.");

        // Optionally report the result.
        Console.WriteLine("Self‑comparison completed successfully. Revision count: " + doc.Revisions.Count);

        // Save the (unchanged) document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SelfComparison.docx");
        doc.Save(outputPath);
    }
}
