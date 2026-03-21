using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Create original document
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello world");

        // Create edited document with extra whitespace
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.Writeln("Hello   world");

        // Configure comparison options to ignore whitespace/formatting changes
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,
            CompareMoves = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the result document containing revisions
        docOriginal.Save("ComparisonResult.docx");
    }
}
