using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It has two lines.");

        // Create the revised document with modifications.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the edited paragraph."); // changed text
        builderRevised.Writeln("It has two lines."); // same line
        builderRevised.Writeln("An additional line in the revised version."); // new line

        // Configure compare options to use the revised document as the target.
        CompareOptions compareOptions = new CompareOptions
        {
            // Setting Target to New makes the other document (original) the base,
            // so revisions will be recorded in the document we call Compare on (revised).
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will appear in the 'revised' document.
        revised.Compare(original, "Author", DateTime.Now, compareOptions);

        // Verify that revisions were generated.
        if (revised.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions in the revised document, but none were found.");

        // Save both documents. The revised document now contains the tracked changes.
        original.Save("Original.docx");
        revised.Save("Revised_With_Revisions.docx");
    }
}
