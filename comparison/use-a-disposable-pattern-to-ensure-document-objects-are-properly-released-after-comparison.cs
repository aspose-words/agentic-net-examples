using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");

        // Create the revised document.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world.");

        // Compare the documents. Provide author name and current date/time.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Accept all revisions.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should be accepted.");
        }

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "compared.docx");
        original.Save(outputPath);

        // Explicitly release references (optional, helps GC).
        builder1 = null;
        builder2 = null;
        original = null;
        revised = null;
    }
}
