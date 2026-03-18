using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceIgnoreFieldsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add normal text that we want to replace.
        builder.Writeln("Hello world!");

        // Insert a field whose result also contains the word "Hello".
        // The field result will be ignored when we set IgnoreFields = true.
        builder.InsertField("QUOTE", "Hello again!");

        // Configure find/replace options to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true   // Skip any text that is part of a field.
        };

        // Replace the word "Hello" with "Greetings" only in the main document body.
        doc.Range.Replace("Hello", "Greetings", options);

        // Update fields so that any field codes are recalculated (if needed).
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
