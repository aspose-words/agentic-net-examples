using System;
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

        // Sample text containing a custom delimiter ';' surrounded by whitespace.
        builder.Writeln("Apple ; Banana ;Cherry ;  Date");

        // Regular expression that matches a semicolon only when it has whitespace on both sides.
        // The look‑behind (?<=\s) ensures a whitespace character precedes the semicolon,
        // and the look‑ahead (?=\s) ensures a whitespace character follows it.
        Regex delimiterRegex = new Regex(@"(?<=\s);(?=\s)");

        // Perform the replacement: replace the matched semicolon with a comma.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(delimiterRegex, ",", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one delimiter replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
