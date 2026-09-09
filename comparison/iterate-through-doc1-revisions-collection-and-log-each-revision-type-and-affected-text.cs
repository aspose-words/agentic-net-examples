using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class RevisionLogger
{
    public static void Main()
    {
        // Create the original document with some text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");
        builderOriginal.Writeln("This line will stay the same.");

        // Create the revised document with modifications.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words!"); // Modified line.
        builderRevised.Writeln("This line will stay the same."); // Unchanged line.
        builderRevised.Writeln("An extra line added."); // New line.

        // Compare the documents to generate revisions in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were created.
        if (original.Revisions.Count == 0)
        {
            Console.WriteLine("No revisions were detected.");
        }
        else
        {
            // Iterate through each revision and log its type and affected text.
            foreach (Revision rev in original.Revisions)
            {
                string affectedText = rev.ParentNode?.GetText().Trim() ?? string.Empty;
                Console.WriteLine($"Revision type: {rev.RevisionType}, affected text: \"{affectedText}\"");
            }
        }

        // Save the compared document (optional artifact).
        original.Save("Compared.docx");
    }
}
