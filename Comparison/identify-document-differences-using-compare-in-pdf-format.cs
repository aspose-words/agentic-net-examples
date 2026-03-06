using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Documents must not contain revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Save the document with revisions (differences) as a PDF file.
        docOriginal.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }
}
