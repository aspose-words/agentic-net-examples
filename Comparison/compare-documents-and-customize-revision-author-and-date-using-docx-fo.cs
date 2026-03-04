using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparisonExample
{
    static void Main()
    {
        // Create the original document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is the original paragraph.");
        builder.Writeln("Second line of original document.");

        // Clone the original document to create an edited version.
        Document editedDoc = (Document)originalDoc.Clone(true);
        // Modify the edited document.
        editedDoc.FirstSection.Body.FirstParagraph.Runs[0].Text = "This is the edited paragraph.";
        // Add an extra line to the edited document.
        DocumentBuilder editedBuilder = new DocumentBuilder(editedDoc);
        editedBuilder.Writeln("Additional line in edited document.");

        // Ensure both documents have no revisions before comparison.
        if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
        {
            // Compare the documents, specifying custom author initials and revision date.
            string revisionAuthor = "JD"; // Author initials.
            DateTime revisionDate = new DateTime(2023, 12, 31, 15, 30, 0); // Custom date and time.
            originalDoc.Compare(editedDoc, revisionAuthor, revisionDate);
        }

        // Save the result (original document now contains revisions) to a DOCX file.
        originalDoc.Save("ComparedDocument.docx");
    }
}
