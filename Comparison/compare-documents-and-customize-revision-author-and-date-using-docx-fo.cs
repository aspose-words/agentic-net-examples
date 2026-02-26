using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Create the original document and add some content.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("This is the original document.");

        // Create the edited document and add modified content.
        Document docEdited = new Document();
        builder = new DocumentBuilder(docEdited);
        builder.Writeln("This is the edited document.");

        // Ensure both documents have no revisions before performing the comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Define custom author initials and revision date.
            string author = "JD"; // Author initials.
            DateTime revisionDate = new DateTime(2023, 12, 31, 15, 30, 0); // Custom date and time.

            // Compare the documents. The resulting revisions will carry the specified author and date.
            docOriginal.Compare(docEdited, author, revisionDate);
        }

        // Save the document that now contains the revisions.
        docOriginal.Save("ComparedResult.docx");
    }
}
