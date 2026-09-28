using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        // Some body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter.");

        // Second heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        // Some body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the second chapter.");

        // Perform find-and-replace: insert a page break after each heading.
        // The metacharacter \f inserts a page break. $0 keeps the original matched text.
        Regex headingRegex = new Regex(@"Chapter \d+");
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(headingRegex, "$0\\f", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }
}
