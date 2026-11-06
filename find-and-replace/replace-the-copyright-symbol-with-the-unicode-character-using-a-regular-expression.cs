using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph containing the text "(c)".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document containing the copyright placeholder (c).");

        // Define a regular expression that matches the literal "(c)".
        Regex copyrightRegex = new Regex(@"\(c\)", RegexOptions.Compiled);

        // Replace all occurrences of the pattern with the Unicode © character.
        int replacedCount = doc.Range.Replace(copyrightRegex, "©", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
