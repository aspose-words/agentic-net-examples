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

        // Add sample text containing the custom delimiter "||".
        builder.Writeln("Alpha || Beta");
        builder.Writeln("Gamma||Delta");
        builder.Writeln("Epsilon   ||   Zeta");
        builder.Writeln("NoDelimiterHere");

        // Define a regex that captures any whitespace before and after the delimiter.
        // The delimiter is "||". The captured groups keep the surrounding whitespace.
        Regex delimiterRegex = new Regex(@"(?<pre>\s*)\|\|(?<post>\s*)");

        // Replace the delimiter with a comma, preserving the captured whitespace.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(delimiterRegex, "${pre},${post}", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one delimiter replacement.");

        // Save the modified document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
