using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the original, edited, and output documents.
        string originalPath = "Original.docx";
        string editedPath = "Edited.docx";
        string outputPath = "Compared.docx";

        // Load the two documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited = new Document(editedPath);

        // Remove any existing revisions from both documents before comparison.
        if (docOriginal.Revisions.Count > 0)
            docOriginal.Revisions.AcceptAll(); // or RejectAll() depending on desired behavior

        if (docEdited.Revisions.Count > 0)
            docEdited.Revisions.AcceptAll();

        // Perform the comparison. The original document will receive the revisions.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the resulting document with revisions.
        docOriginal.Save(outputPath);
    }
}
