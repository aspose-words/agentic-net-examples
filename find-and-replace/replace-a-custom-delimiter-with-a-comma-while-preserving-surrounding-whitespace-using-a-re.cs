using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First value || second value.");
        builder.Writeln("Third value||fourth value.");
        builder.Writeln("No delimiter here.");

        // Save the original for reference (optional).
        doc.Save("input.docx");

        // Define a regex that matches the custom delimiter "||".
        Regex delimiterRegex = new Regex(@"\|\|");

        // Perform the replacement: replace "||" with a comma while leaving surrounding whitespace untouched.
        int replacedCount = doc.Range.Replace(delimiterRegex, ",", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one delimiter replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
