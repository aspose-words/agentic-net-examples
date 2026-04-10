using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some text.
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("Hello World!");

        // Create the edited document where the text differs only by case.
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.Writeln("hello world!");

        // Configure comparison options to ignore case changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreCaseChanges = true,
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. The author name and date are required.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // The comparison should produce zero revisions because case differences are ignored.
        int revisionCount = docOriginal.Revisions.Count;
        Console.WriteLine($"Number of revisions after comparison: {revisionCount}");

        // Save the resulting document (it will contain no revisions in this case).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        docOriginal.Save(outputPath);
        Console.WriteLine($"Comparison result saved to: {outputPath}");
    }
}
