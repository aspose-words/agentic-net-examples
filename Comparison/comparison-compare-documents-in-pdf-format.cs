using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original PDF document.
        Document docOriginal = new Document("Original.pdf");

        // Load the edited PDF document.
        Document docEdited = new Document("Edited.pdf");

        // Compare the documents. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the comparison result (with revisions) as a PDF file.
        docOriginal.Save("ComparisonResult.pdf");
    }
}
