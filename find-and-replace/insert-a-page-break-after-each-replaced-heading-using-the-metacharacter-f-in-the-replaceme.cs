using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.Writeln("Content under heading 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 2");
        builder.Writeln("Content under heading 2.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 3");
        builder.Writeln("Content under heading 3.");

        // Save the original document (optional, for inspection).
        string inputPath = Path.Combine(Environment.CurrentDirectory, "Input.docx");
        doc.Save(inputPath);

        // Replace each heading with the same text followed by a page break using the \f metacharacter.
        Regex headingRegex = new Regex(@"Heading \d+");
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true // Enable $0 substitution.
        };
        int replacedCount = doc.Range.Replace(headingRegex, "$0\f", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
        doc.Save(outputPath);
    }
}
