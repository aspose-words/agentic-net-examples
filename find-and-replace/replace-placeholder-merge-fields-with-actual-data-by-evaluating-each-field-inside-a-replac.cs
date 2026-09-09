using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required for Aspose.Words font/color types.

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholder merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear <<FirstName>> <<LastName>>,");
        builder.Writeln("Your order <<OrderId>> has been shipped from <<Company>>.");
        builder.Writeln("Thank you!");

        // Data that will replace the placeholders.
        var data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "OrderId", "12345" },
            { "Company", "Acme Corp" }
        };

        // Set up the find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PlaceholderReplacer(data);

        // Regex that matches placeholders of the form <<Placeholder>>.
        Regex placeholderPattern = new Regex(@"<<(\w+)>>", RegexOptions.Compiled);

        // Perform the replace operation. The replacement string is ignored because the callback
        // supplies the actual replacement text.
        int replacedCount = doc.Range.Replace(placeholderPattern, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Implements IReplacingCallback to supply replacement text based on the captured placeholder name.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _data;

        public PlaceholderReplacer(IDictionary<string, string> data)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Group 1 contains the placeholder name without the surrounding << >>.
            string key = args.Match.Groups[1].Value;

            if (_data.TryGetValue(key, out string value))
                args.Replacement = value; // Replace with the value from the dictionary.
            else
                args.Replacement = args.Match.Value; // Keep the original placeholder if not found.

            return ReplaceAction.Replace;
        }
    }
}
