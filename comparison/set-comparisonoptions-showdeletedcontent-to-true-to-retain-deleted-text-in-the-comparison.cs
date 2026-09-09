using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with two paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Paragraph one.");
        builderOriginal.Writeln("Paragraph two – this will be deleted.");

        // Create the revised document that lacks the second paragraph.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Paragraph one.");

        // Configure comparison options so that deleted content is retained in the result.
        // Setting the target to the new document makes deletions appear in the comparison output.
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions (including deletions) will be added to 'original'.
        original.Compare(revised, "John Doe", DateTime.Now, compareOptions);

        // Verify that at least one revision (the deletion) was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Save the comparison result to a file.
        original.Save("ComparisonResult.docx");
    }
}
