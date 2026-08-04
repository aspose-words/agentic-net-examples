using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample content containing the custom tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(@"Here is a tag: <customTag attr=""oldValue1"">Content</customTag>");
        builder.Writeln(@"Another tag: <customTag attr=""oldValue2"">More content</customTag>");
        builder.Writeln(@"A tag without the attribute should stay unchanged: <customTag>Nothing</customTag>");

        // Define a regular expression that captures the attribute value.
        // Group 1 = the part before the value, Group 2 = the closing quote.
        Regex regex = new Regex(@"(<customTag\s+attr="")[^""]+("")", RegexOptions.IgnoreCase);

        // Replacement uses the captured groups and inserts the new attribute value.
        const string newValue = "newValue";
        string replacement = $"$1{newValue}$2";

        // Enable substitution so that $1 and $2 are replaced with the captured groups.
        FindReplaceOptions options = new FindReplaceOptions { UseSubstitutions = true };

        // Perform the replacement across the whole document.
        int replacedCount = doc.Range.Replace(regex, replacement, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No attribute values were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Verify the result by reading the saved file.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        // Simple check to ensure the new value appears.
        if (!resultText.Contains(newValue))
            throw new InvalidOperationException("Replacement verification failed.");
    }
}
