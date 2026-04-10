using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original paragraph.");
        builder.Writeln("It contains some text that will be changed.");

        // Create the edited document with intentional differences.
        Document edited = new Document();
        builder = new DocumentBuilder(edited);
        builder.Writeln("This is the edited paragraph."); // Changed text.
        builder.Writeln("It contains some text that has been modified."); // Changed text.

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");
        }

        // Perform the comparison. The original document will receive revision marks.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Comparison did not produce any revisions.");
        }

        // Save the compared document preserving all revision metadata.
        string resultPath = Path.Combine(outputDir, "ComparedDocument.docx");
        original.Save(resultPath);
    }
}
