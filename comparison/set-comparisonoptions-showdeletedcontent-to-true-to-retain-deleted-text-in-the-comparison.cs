using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Paragraph 1.");
        builder.Writeln("Paragraph to be deleted.");
        builder.Writeln("Paragraph 3.");

        // Clone the original document to serve as the edited version.
        Document docEdited = (Document)docOriginal.Clone(true);

        // Remove the second paragraph from the edited document to simulate a deletion.
        docEdited.FirstSection.Body.Paragraphs[1].Remove();

        // Configure comparison options (no special flags are required for showing deletions).
        CompareOptions compareOptions = new CompareOptions();

        // Compare the original document with the edited one.
        docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);

        // Save the comparison result.
        docOriginal.Save("ComparisonResult.docx");
    }
}
