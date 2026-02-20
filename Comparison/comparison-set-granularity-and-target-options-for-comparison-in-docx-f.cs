using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class ComparisonExample
{
    static void Main()
    {
        // Create the original document and add some content.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Clone the original document to simulate an edited version.
        Document docEdited = (Document)docOriginal.Clone(true);
        // Make a simple edit in the cloned document.
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "The quick brown cat jumps over the lazy dog.";

        // Configure comparison options:
        // - Track changes at the character level.
        // - Use the edited document as the target (i.e., compare against the new version).
        CompareOptions compareOptions = new CompareOptions
        {
            Granularity = Granularity.CharLevel,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Save the result as a DOCX file using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        // (Optional) Set compliance level if needed; default is Ecma376_2006.
        // saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        docOriginal.Save("ComparisonResult.docx", saveOptions);
    }
}
