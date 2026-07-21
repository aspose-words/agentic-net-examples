using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Sample data for placeholders.
        var placeholderData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "Email", "john.doe@example.com" }
        };

        // 1. Create a sample document containing placeholder merge fields.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Dear {{FirstName}} {{LastName}},");
        builder.Writeln("Your registered email is {{Email}}.");
        builder.Writeln("Thank you for using our service.");
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // 2. Load the document we just created.
        var doc = new Document(templatePath);

        // 3. Define a regex that matches placeholders like {{FieldName}}.
        var placeholderRegex = new Regex(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

        // 4. Perform the replace using a custom IReplacingCallback implementation.
        var options = new FindReplaceOptions(new PlaceholderReplacer(placeholderData));
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced. Check the regex pattern and data.");

        // 5. Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Callback that replaces each matched placeholder with the corresponding value from the dictionary.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _data;

        public PlaceholderReplacer(IDictionary<string, string> data)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The first captured group contains the field name without braces.
            string fieldName = args.Match.Groups[1].Value;

            if (_data.TryGetValue(fieldName, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = args.Match.Value; // keep original placeholder if not found

            return ReplaceAction.Replace;
        }
    }
}
