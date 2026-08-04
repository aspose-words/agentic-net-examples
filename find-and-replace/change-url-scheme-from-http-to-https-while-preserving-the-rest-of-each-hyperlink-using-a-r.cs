using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a few HTTP hyperlinks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample hyperlinks:");
        builder.InsertHyperlink("Aspose", "http://www.aspose.com", false);
        builder.Writeln();
        builder.InsertHyperlink("GitHub", "http://github.com", false);
        builder.Writeln();
        builder.InsertHyperlink("StackOverflow", "http://stackoverflow.com", false);
        builder.Writeln();

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Replace the HTTP scheme with HTTPS using a regular expression.
        Regex httpPattern = new Regex(@"http://", RegexOptions.IgnoreCase);
        int replacedCount = loaded.Range.Replace(httpPattern, "https://", new FindReplaceOptions());

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTTP URLs were found to replace.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
