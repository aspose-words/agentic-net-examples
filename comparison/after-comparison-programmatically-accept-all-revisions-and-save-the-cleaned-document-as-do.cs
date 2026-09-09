using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It has two lines.");

        // Create the revised document with some changes.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the edited paragraph."); // changed text
        builderRevised.Writeln("It has two lines."); // same line
        builderRevised.Writeln("An additional line was added."); // new line

        // Compare the documents. The comparison adds revisions to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");

        // Accept all revisions, turning the original document into the revised version.
        original.Revisions.AcceptAll();

        // Ensure all revisions have been accepted.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Save the cleaned document.
        string outputPath = "CleanedDocument.docx";
        original.Save(outputPath);
    }
}
