using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a custom tag that contains an attribute.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<customTag attr=\"oldValue\">Sample content</customTag>");
        builder.Writeln("<customTag attr=\"anotherOldValue\">More content</customTag>");

        // Save the sample document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define a regular expression that matches the attribute value of the custom tag.
        // This pattern matches attr="any characters except a quote".
        Regex regex = new Regex(@"attr=""[^""]*""", RegexOptions.Compiled);

        // Replace the matched attribute value with the new desired value.
        const string replacement = @"attr=""newValue""";
        int replacedCount = loadedDoc.Range.Replace(regex, replacement, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one attribute replacement, but none were made.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
