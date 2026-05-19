using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few headings (using the built‑in Heading1 style) and some body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the second chapter.");

        // Save the original document (optional, just for reference).
        doc.Save("input.docx");

        // Prepare a regular expression that matches the whole heading text.
        Regex headingRegex = new Regex(@"Chapter \d+");

        // Configure find‑replace options to enable substitution patterns.
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true   // Allows $0 to represent the whole match.
        };

        // Replace each heading with itself followed by a page break.
        // The form‑feed character (\f) is interpreted by Aspose.Words as a page break.
        int replacedCount = doc.Range.Replace(headingRegex, "$0\f", options);

        // Verify that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
