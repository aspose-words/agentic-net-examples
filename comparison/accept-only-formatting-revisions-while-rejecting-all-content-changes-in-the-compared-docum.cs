using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Hello World!"); // Plain text.

        // Create the edited document with both content and formatting changes.
        Document edited = new Document();
        DocumentBuilder editedBuilder = new DocumentBuilder(edited);
        // Change the text content.
        editedBuilder.Writeln("Hello Aspose!");
        // Apply a formatting change (make the text italic).
        editedBuilder.Font.Italic = true;
        editedBuilder.Writeln("Formatted line.");

        // Compare the documents, generating revisions for both content and formatting changes.
        // No special compare options are needed because we want to capture formatting revisions.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Process revisions: accept only formatting revisions, reject all others.
        // Create a snapshot of the revisions to avoid modifying the collection while iterating.
        var revisions = original.Revisions.Cast<Revision>().ToList();

        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();   // Keep formatting changes.
            else
                rev.Reject();   // Discard content changes.
        }

        // Save the resulting document.
        string resultPath = Path.Combine(outputDir, "Result.docx");
        original.Save(resultPath);
    }
}
