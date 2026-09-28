using System;
using System.IO;
using Aspose.Words;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It contains several lines of text.");

        // Create the revised document with differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised paragraph."); // Modified line.
        builderRevised.Writeln("It contains several lines of text."); // Same line.
        builderRevised.Writeln("An additional line is added."); // New line.

        // Perform comparison. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Accept all revisions, cleaning the document.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should have been accepted.");
        }

        // Save the cleaned document as DOCX.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "cleaned.docx");
        original.Save(outputPath, SaveFormat.Docx);
    }
}
