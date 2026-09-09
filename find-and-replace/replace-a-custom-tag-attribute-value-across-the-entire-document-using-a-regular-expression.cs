using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the rule set
using Newtonsoft.Json; // Required by the rule set

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample content that contains the custom tag with varying attribute values.
        builder.Writeln(@"Here is a custom tag: <custom attr=""value1"">Some text</custom>");
        builder.Writeln(@"Another occurrence: <custom attr=""oldValue"">More text</custom>");
        builder.Writeln(@"And one more: <custom attr=""value2""/>");

        // Define a regular expression that captures the attribute value of the <custom> tag.
        // Group 1: the opening part up to the attribute value quote.
        // Group 2: the attribute value itself (to be replaced).
        // Group 3: the closing quote and the rest of the tag.
        Regex regex = new Regex(@"(<custom\s+attr="")([^""]+)(""[^>]*>)", RegexOptions.Compiled);

        // Set up find/replace options to enable substitution groups in the replacement pattern.
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true,
            LegacyMode = false
        };

        // Replacement string uses the captured groups, inserting the new attribute value.
        string replacement = "$1newValue$3";

        // Perform the replacement across the whole document.
        int replacedCount = doc.Range.Replace(regex, replacement, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one attribute value replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: write a simple JSON report about the operation (demonstrates required package usage).
        var report = new
        {
            ReplacementsMade = replacedCount,
            OutputFile = outputPath
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        System.IO.File.WriteAllText("report.json", jsonReport);
    }
}
