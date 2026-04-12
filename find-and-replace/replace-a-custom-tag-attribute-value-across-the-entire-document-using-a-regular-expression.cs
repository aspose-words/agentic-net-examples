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

        // Insert sample XML‑like tags that contain an attribute we want to change.
        builder.Writeln("<customTag attr=\"oldValue1\">Content1</customTag>");
        builder.Writeln("<customTag attr=\"oldValue2\">Content2</customTag>");
        builder.Writeln("<customTag attr=\"oldValue3\">Content3</customTag>");

        // Regular expression that matches the attribute value.
        // Captures three groups: the prefix (attr="), the current value, and the closing quote.
        Regex regex = new Regex("(attr=\")([^\"]+)(\")", RegexOptions.IgnoreCase);

        // Replacement string uses substitution syntax to keep the first and third groups
        // and replace the second group (the attribute value) with the new value.
        string replacement = "$1newValue$3";

        // Configure find/replace options to enable substitution handling.
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true,
            LegacyMode = false
        };

        // Perform the replacement across the entire document.
        int replacedCount = doc.Range.Replace(regex, replacement, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No attribute values were replaced.");

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
