using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains custom tags with an attribute.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(@"Here is a custom tag: <customTag attr=""oldValue"">Content</customTag>");
        builder.Writeln(@"Another line with the same attribute: <customTag attr=""oldValue"">More</customTag>");
        builder.Writeln(@"A line without the tag should stay unchanged.");

        // Save the sample document so that the later load step works on a real file.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and replace the attribute value using a regex.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regex matches the opening part of the attribute and captures the surrounding quotes.
        // Group 1: <customTag ... attr="
        // Group 2: closing quote (")
        Regex regex = new Regex(@"(<customTag\s+[^>]*attr="")[^""]*("")", RegexOptions.IgnoreCase);

        FindReplaceOptions options = new FindReplaceOptions
        {
            // Enable substitution groups ($1, $2) in the replacement string.
            UseSubstitutions = true
        };

        // Replace the captured value with "newValue".
        int replacedCount = loaded.Range.Replace(regex, "$1newValue$2", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one attribute replacement, but none were made.");

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // Optional: Verify the replacement by reading back the text.
        Document verify = new Document(outputPath);
        string text = verify.GetText();

        // Simple sanity check – the new value should be present.
        if (!text.Contains(@"attr=""newValue"""))
            throw new InvalidOperationException("The attribute value was not replaced as expected.");

        // The program finishes without requiring any user interaction.
    }
}
