using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareGranularityExample
{
    public static void Run()
    {
        // Create the original document and add some text.
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

        // Create the edited document and add modified text.
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.Writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the character level.
            Granularity = Granularity.CharLevel,
            // Use the edited document as the target (equivalent to "Show changes in: New document").
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the comparison result in DOCX format.
        docOriginal.Save("ComparisonResult.docx");
    }
}

class Program
{
    static void Main(string[] args)
    {
        CompareGranularityExample.Run();
    }
}
