using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // ------------------------------------------------------------
        // Replace text that appears inside fields.
        // ------------------------------------------------------------
        // By default fields are included, but we set the option explicitly
        // for clarity.
        FindReplaceOptions replaceInFields = new FindReplaceOptions
        {
            IgnoreFields = false
        };
        int replacedInFields = doc.Range.Replace("Hello", "Greetings", replaceInFields);
        Console.WriteLine($"Replacements performed inside fields: {replacedInFields}");

        // ------------------------------------------------------------
        // Replace text that appears only outside of fields.
        // ------------------------------------------------------------
        // Setting IgnoreFields to true tells the engine to skip whole fields.
        FindReplaceOptions replaceOutsideFields = new FindReplaceOptions
        {
            IgnoreFields = true
        };
        int replacedOutsideFields = doc.Range.Replace("World", "Universe", replaceOutsideFields);
        Console.WriteLine($"Replacements performed outside fields: {replacedOutsideFields}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
