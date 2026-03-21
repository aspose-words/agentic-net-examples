using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Create the original document.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello world!");
        builder.Writeln("This paragraph will be deleted.");

        // Clone the original and delete the second paragraph without tracking revisions.
        Document docEdited = (Document)docOriginal.Clone(true);
        docEdited.FirstSection.Body.Paragraphs[1].Remove();

        // Set up comparison options (ShowDeletedContent is not available in this version).
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison; revisions (including deletions) are added to docOriginal.
        docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);

        // Save the comparison result.
        docOriginal.Save("ComparisonResult.docx");
    }
}
