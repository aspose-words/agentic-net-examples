using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new document and add some sample text containing numbers in US format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sales figures are 1,234.56 for Q1 and 12,345.78 for Q2.");

        // Define the regular expression that matches numbers formatted with commas as thousand separators
        // and a dot as the decimal separator (e.g., 1,234.56 or 12,345).
        Regex numberPattern = new Regex(@"\d{1,3}(?:,\d{3})*(?:\.\d+)?");

        // Set up FindReplaceOptions with a custom callback that will format each match
        // according to the desired locale (French in this example).
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new LocalizedNumberReplacer(new CultureInfo("fr-FR"));

        // Perform the replace operation. The replacement string is ignored because the callback
        // supplies the actual replacement value.
        doc.Range.Replace(numberPattern, string.Empty, options);

        // Save the modified document.
        doc.Save("LocalizedNumbers.docx");
    }
}

// Implements IReplacingCallback to convert each matched number to a localized format.
class LocalizedNumberReplacer : IReplacingCallback
{
    private readonly CultureInfo _targetCulture;

    public LocalizedNumberReplacer(CultureInfo targetCulture)
    {
        _targetCulture = targetCulture;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Parse the matched number using invariant culture (which matches the US format).
        if (double.TryParse(args.Match.Value, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint,
                            CultureInfo.InvariantCulture, out double number))
        {
            // Format the number using the target culture's number format.
            // "N" includes group separators and decimal digits.
            args.Replacement = number.ToString("N", _targetCulture);
        }
        else
        {
            // If parsing fails, keep the original text.
            args.Replacement = args.Match.Value;
        }

        return ReplaceAction.Replace;
    }
}
