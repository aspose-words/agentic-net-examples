using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // 1. Replace text that appears inside fields.
        //    The default FindReplaceOptions does NOT ignore fields,
        //    so a simple call to Range.Replace will affect field contents.
        // -----------------------------------------------------------------
        // Example: replace the placeholder _FullName_ with an actual name.
        int replacedInFields = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Replacements inside fields: {replacedInFields}");

        // -----------------------------------------------------------------
        // 2. Replace text while ignoring everything that is inside fields.
        //    Set FindReplaceOptions.IgnoreFields = true to skip field text.
        // -----------------------------------------------------------------
        FindReplaceOptions ignoreFieldsOptions = new FindReplaceOptions
        {
            IgnoreFields = true
        };
        // Example: replace the word "Hello" only in normal document text.
        int replacedOutsideFields = doc.Range.Replace("Hello", "Greetings", ignoreFieldsOptions);
        Console.WriteLine($"Replacements outside fields: {replacedOutsideFields}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
