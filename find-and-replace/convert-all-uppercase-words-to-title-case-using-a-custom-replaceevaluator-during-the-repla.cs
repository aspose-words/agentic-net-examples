using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample text containing uppercase words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS is a TEST of the UPPERCASE WORDS like NASA and ASP.NET.");
        builder.Writeln("Another LINE with MIXED case WORDS such as HELLO and WORLD.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new UppercaseToTitleCaseReplacer()
        };

        // Regex matches whole words composed of two or more uppercase letters.
        Regex uppercasePattern = new Regex(@"\b[A-Z]{2,}\b");

        // Perform the replace operation. The callback will supply the actual replacement text.
        int replacementCount = doc.Range.Replace(uppercasePattern, string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No uppercase words were found to replace.");

        // Save the modified document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }

    // Callback that converts each matched uppercase word to title case.
    private class UppercaseToTitleCaseReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string original = args.Match.Value;
            // Convert to title case: first letter uppercase, the rest lowercase.
            string titleCase = char.ToUpper(original[0]) + original.Substring(1).ToLowerInvariant();

            args.Replacement = titleCase;
            return ReplaceAction.Replace;
        }
    }
}
