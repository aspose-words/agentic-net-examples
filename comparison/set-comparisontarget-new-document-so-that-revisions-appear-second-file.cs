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
        builder.Writeln("Hello world! This is the original paragraph.");

        // Clone the original and make a change – this will be the second document.
        Document docEdited = (Document)docOriginal.Clone(true);
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "Hello world! This is the edited paragraph.";

        // Configure comparison so that the base document is the original (Target = New).
        // Revisions will therefore be recorded in the document on which Compare is called (docEdited).
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison; revisions are added to docEdited.
        docEdited.Compare(docOriginal, "John Doe", DateTime.Now, compareOptions);

        // Save both files – the edited file now contains the revision marks.
        docOriginal.Save("Original.docx");
        docEdited.Save("Edited_With_Revisions.docx");
    }
}
