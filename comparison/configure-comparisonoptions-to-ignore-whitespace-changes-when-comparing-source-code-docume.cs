using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original documentation document.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("public void Foo()");
        builder.Writeln("{");
        builder.Writeln("    // This is a comment");
        builder.Writeln("    Console.WriteLine(\"Hello\");");
        builder.Writeln("}");

        // Clone the original and introduce whitespace changes.
        Document docEdited = (Document)docOriginal.Clone(true);
        // Add extra spaces inside the method signature.
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        if (firstParagraph.Runs.Count > 0)
            firstParagraph.Runs[0].Text = "public   void   Foo()";

        // Insert an empty line after the opening brace.
        Paragraph emptyParagraph = new Paragraph(docEdited);
        docEdited.FirstSection.Body.InsertAfter(emptyParagraph, docEdited.FirstSection.Body.Paragraphs[1]);

        // Configure comparison options to ignore formatting (which includes whitespace changes).
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = true,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Output the number of revisions detected.
        int revisionCount = docOriginal.Revisions.Count;
        Console.WriteLine($"Revisions count after ignoring whitespace: {revisionCount}");

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        docOriginal.Save(outputPath);
    }
}
