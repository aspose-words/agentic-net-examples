using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class CustomTagAttributeReplacer
{
    // Replaces the value of a custom XML attribute (or any text matching the pattern)
    // throughout the whole document using a regular expression.
    public static void ReplaceAttributeValue(string inputFilePath, string outputFilePath,
                                             string attributeName, string newValue)
    {
        // Load the document.
        Document doc = new Document(inputFilePath);

        // Build a regex that finds the attribute name followed by an equals sign and a quoted value.
        // Example pattern for attributeName = "oldValue":  attributeName\s*=\s*".*?"
        string pattern = $@"{Regex.Escape(attributeName)}\s*=\s*""[^""]*""";

        // Replacement string keeps the attribute name and inserts the new quoted value.
        string replacement = $"{attributeName}=\"{newValue}\"";

        // Use FindReplaceOptions to treat StructuredDocumentTag content as plain text,
        // otherwise the replacement would not cross tag boundaries.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreStructuredDocumentTags = true   // treat SDT content as simple text
        };

        // Perform the replace operation on the whole document range.
        doc.Range.Replace(new Regex(pattern), replacement, options);

        // Save the modified document.
        doc.Save(outputFilePath);
    }

    // Example usage.
    public static void Main()
    {
        // Create a temporary source document with sample content.
        string tempDir = Path.GetTempPath();
        string inputPath = Path.Combine(tempDir, "Source.docx");
        string outputPath = Path.Combine(tempDir, "Result.docx");

        // Build a simple document containing the attribute to be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(@"This is a test document. myAttr=""OldValue"" appears here.");
        doc.Save(inputPath);

        // Replace the value of the custom attribute "myAttr" with "NewValue".
        ReplaceAttributeValue(inputPath, outputPath, "myAttr", "NewValue");

        Console.WriteLine($"Attribute value replacement completed. Output saved to: {outputPath}");
    }
}
