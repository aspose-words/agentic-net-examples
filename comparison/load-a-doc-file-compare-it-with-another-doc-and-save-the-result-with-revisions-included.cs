using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");
        builder.Writeln("Paragraph that will be changed in the edited version.");

        // Create the edited document that differs from the original.
        Document edited = new Document();
        DocumentBuilder editedBuilder = new DocumentBuilder(edited);
        editedBuilder.Writeln("This is the edited document.");
        editedBuilder.Writeln("Paragraph that has been changed in the edited version.");

        // Perform the comparison only when both documents have no existing revisions.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            original.Compare(edited, "John Doe", DateTime.Now);
        }

        // Verify that revisions were generated.
        int revisionCount = original.Revisions.Count;
        Console.WriteLine($"Revisions detected: {revisionCount}");

        // Save the original document now containing the revision markup.
        const string outputFile = "ComparisonResult.docx";
        original.Save(outputFile);
    }
}
