using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class CompareDocuments
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Set up comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Define custom author and date for the revisions.
        string customAuthor = "Alice Smith";
        DateTime customDate = new DateTime(2023, 12, 31, 10, 30, 0);

        // Perform the comparison; revisions will be created with the specified author and date.
        docOriginal.Compare(docEdited, customAuthor, customDate, compareOptions);

        // Ensure all revisions have the desired author and date (in case defaults were used).
        foreach (Revision rev in docOriginal.Revisions)
        {
            rev.Author = customAuthor;
            rev.DateTime = customDate;
        }

        // Save the resulting document as DOCX using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        docOriginal.Save("ComparedResult.docx", saveOptions);
    }
}
