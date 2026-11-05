using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing custom tag attributes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<customTag attr=\"oldValue1\">Content 1</customTag>");
        builder.Writeln("<customTag attr=\"oldValue2\">Content 2</customTag>");
        builder.Writeln("Some other text without the tag.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regex to match the attribute value of <customTag attr="...">.
        // Captures the part before the value, the value itself, and the trailing quote.
        Regex regex = new Regex(@"(<customTag\s+[^>]*attr="")([^""]*)("")", RegexOptions.IgnoreCase);

        // Replacement keeps the surrounding groups and substitutes a new value.
        string replacement = "$1newValue$3";

        // Perform the replacement across the entire document.
        int replacedCount = loaded.Range.Replace(regex, replacement, new FindReplaceOptions());

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No attribute values were replaced.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
