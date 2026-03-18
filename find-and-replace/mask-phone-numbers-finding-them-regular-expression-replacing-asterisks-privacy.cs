using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class PhoneNumberMasking
{
    // Callback that replaces each found phone number with a string of asterisks of the same length.
    private class MaskingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Preserve the original length (including separators) by using asterisks.
            args.Replacement = new string('*', args.Match.Value.Length);
            return ReplaceAction.Replace;
        }
    }

    static void Main()
    {
        // Load an existing document (replace with your actual file path).
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Regular expression that matches common US phone number formats:
        // 123-456-7890, (123) 456-7890, 123 456 7890, 123.456.7890, 1234567890, etc.
        string phonePattern = @"\b(?:\+?1[\s-\.]?)?(?:\(\d{3}\)|\d{3})[\s\-\.]?\d{3}[\s\-\.]?\d{4}\b";

        // Set up find/replace options with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new MaskingCallback());

        // Perform the replace operation using the regex pattern.
        doc.Range.Replace(new Regex(phonePattern), string.Empty, options);

        // Save the modified document (replace with your desired output path).
        doc.Save(@"C:\Docs\Output.docx");
    }
}
