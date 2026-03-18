using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareWithCustomAuthorAndTimestamp
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Define custom author name and timestamp for the revisions created by the comparison.
        string customAuthor = "Alice Smith";
        DateTime customTimestamp = new DateTime(2023, 5, 1, 10, 30, 0);

        // Perform the comparison. The revisions in docOriginal will be attributed to the custom author and timestamp.
        docOriginal.Compare(docEdited, customAuthor, customTimestamp);

        // Save the resulting document with revisions.
        docOriginal.Save("Compared.docx");
    }
}
