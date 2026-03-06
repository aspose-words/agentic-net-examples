using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparisonExample
{
    static void Main()
    {
        // Create the original document.
        Document docOriginal = new Document();                     // create
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a few lines of text.");

        // Create the edited document.
        Document docEdited = new Document();                       // create
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.Writeln("This is the edited document.");    // changed line
        builderEdited.Writeln("It contains a few lines of text."); // unchanged line
        builderEdited.Writeln("Additional line added.");          // new line

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents, specifying custom author initials and revision date.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);   // compare with author and date
        }

        // Save the resulting document that now contains revision marks.
        docOriginal.Save("ComparedResult.docx");                  // save
    }
}
