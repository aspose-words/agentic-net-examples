using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with custom delimiters (pipe character) and surrounding whitespace.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Item1 | Item2|Item3 |Item4  |  Item5");
        // Save the input document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform find-and-replace.
        Document loaded = new Document(inputPath);

        // Define a regex that matches a pipe delimiter with any surrounding whitespace,
        // capturing the whitespace before and after the pipe.
        Regex delimiterRegex = new Regex(@"(?<pre>\s*)\|(?<post>\s*)");

        // Replacement string uses the captured whitespace groups and inserts a comma.
        string replacement = "${pre},${post}";

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(delimiterRegex, replacement, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one delimiter replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
