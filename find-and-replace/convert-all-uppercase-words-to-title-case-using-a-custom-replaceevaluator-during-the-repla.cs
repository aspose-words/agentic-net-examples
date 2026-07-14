using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with uppercase words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS is a TEST document.");
        builder.Writeln("IT contains UPPERCASE WORDS like ASP.NET and CSHARP.");
        builder.Writeln("Mixed case words such as Hello should stay unchanged.");

        // Save the original document (optional, for inspection).
        doc.Save("input.docx");

        // Set up the replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new UppercaseToTitleCaseReplacer();

        // Regex matches whole words consisting of two or more uppercase letters.
        Regex uppercaseWordPattern = new Regex(@"\b[A-Z]{2,}\b");

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(uppercaseWordPattern, "", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one uppercase word to be replaced.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that converts each matched uppercase word to title case.
    private class UppercaseToTitleCaseReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string original = args.Match.Value;
            string titleCased = ToTitleCase(original);
            args.Replacement = titleCased;
            return ReplaceAction.Replace;
        }

        // Converts a word like "HELLO" to "Hello".
        private static string ToTitleCase(string word)
        {
            if (string.IsNullOrEmpty(word))
                return word;

            if (word.Length == 1)
                return char.ToUpper(word[0]).ToString();

            return char.ToUpper(word[0]) + word.Substring(1).ToLower();
        }
    }
}
