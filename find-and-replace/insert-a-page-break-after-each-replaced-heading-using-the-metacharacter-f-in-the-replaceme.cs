using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.Writeln("Content of chapter 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.Writeln("Content of chapter 2.");

        // Save the source document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find-and-replace.
        Document loaded = new Document(inputPath);

        // Define a regex that matches the whole heading text.
        Regex headingRegex = new Regex(@"Chapter \d+");

        // Replacement string uses the metacharacter \f to insert a page break after the match.
        string replacement = "$0\\f";

        // Perform the replacement using FindReplaceOptions.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(headingRegex, replacement, options);

        // Validate that at least one heading was replaced.
        if (replacedCount == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
