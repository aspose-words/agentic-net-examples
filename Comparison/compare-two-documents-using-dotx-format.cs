using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDotxExample
{
    static void Main()
    {
        // Load the two documents that will be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // The Compare method requires that both documents have no existing revisions.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. Revisions describing the differences are added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the resulting document (which now contains revision marks) as a DOTX template.
        docOriginal.Save("ComparedTemplate.dotx");
    }
}
