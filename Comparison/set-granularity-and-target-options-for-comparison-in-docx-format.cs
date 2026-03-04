using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocumentsExample
{
    static void Main()
    {
        // Create the original document and add some text.
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

        // Create the edited document and add modified text.
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.Writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");

        // Configure comparison options:
        // - Track changes at the character level.
        // - Use the edited document as the target (equivalent to Word's "Show changes in New").
        CompareOptions compareOptions = new CompareOptions
        {
            Granularity = Granularity.CharLevel,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the resulting document that contains the tracked revisions.
        docOriginal.Save("ComparisonResult.docx");
    }
}
