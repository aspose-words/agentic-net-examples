using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class HeaderFooterComparison
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document original = new Document("Original.docx");
        Document revised = new Document("Revised.docx");

        // Configure comparison options to ignore changes in headers and footers.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreHeadersAndFooters = true
        };

        // Perform the comparison. The revisions will be added to the 'original' document.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Save the result with revisions applied.
        original.Save("ComparisonResult.docx");
    }
}
