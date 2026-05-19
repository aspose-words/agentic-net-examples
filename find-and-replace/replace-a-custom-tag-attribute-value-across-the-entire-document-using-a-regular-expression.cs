using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;      // Required package (not used directly)
using Newtonsoft.Json;    // Required package (not used directly)

public class ReplaceCustomTagAttribute
{
    public static void Main()
    {
        // Create a sample document with custom tags that contain an attribute to be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<custom attr=\"oldValue1\"/>");
        builder.Writeln("<custom attr=\"oldValue2\"/>");
        builder.Writeln("<custom attr=\"keep\"/>");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the attribute value of the custom tag.
        // This pattern finds attr="anyValue" and captures the whole attribute.
        Regex attributeRegex = new Regex(@"attr=""[^""]*""", RegexOptions.IgnoreCase);

        // Replace every occurrence of the attribute with a new value.
        string replacement = @"attr=""newValue""";
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(attributeRegex, replacement, options);

        // Validate that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("No attribute values were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacedCount} attribute occurrence(s). Output saved to '{outputPath}'.");
    }
}
