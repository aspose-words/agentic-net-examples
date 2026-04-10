using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DocumentComparisonExample
{
    public static void Main()
    {
        // Determine a folder for output files (the current directory).
        string outputFolder = Directory.GetCurrentDirectory();

        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder originalBuilder = new DocumentBuilder(original);
        originalBuilder.Writeln("This is the original document content.");

        // Create the edited document with a deliberate difference.
        Document edited = new Document();
        DocumentBuilder editedBuilder = new DocumentBuilder(edited);
        editedBuilder.Writeln("This is the edited document content with a change.");

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Perform the comparison. Revisions will be added to the original document.
            original.Compare(edited, "Comparer", DateTime.Now);
        }

        // Verify that at least one revision was created.
        if (original.Revisions.Count > 0)
        {
            Console.WriteLine($"Revisions detected: {original.Revisions.Count}");
        }
        else
        {
            Console.WriteLine("No revisions were detected.");
        }

        // Save the document that now contains the revision markup.
        string resultPath = Path.Combine(outputFolder, "ComparisonResult.docx");
        original.Save(resultPath);
        Console.WriteLine($"Comparison result saved to: {resultPath}");

        // Accept all revisions to transform the original into the edited version.
        original.Revisions.AcceptAll();

        // Save the accepted version.
        string acceptedPath = Path.Combine(outputFolder, "ComparisonAccepted.docx");
        original.Save(acceptedPath);
        Console.WriteLine($"Accepted document saved to: {acceptedPath}");

        // Release references (Document does not implement IDisposable, so we simply null them).
        original = null;
        edited = null;
    }
}
