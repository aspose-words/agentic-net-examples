using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options to ignore all formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true   // Do not track formatting differences.
        };

        // Perform the comparison. The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the result document which now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
