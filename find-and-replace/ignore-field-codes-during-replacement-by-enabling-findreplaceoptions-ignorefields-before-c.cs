using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!"); // This occurrence should be replaced.

        // Insert a field that also contains the word "Hello".
        // The field result will be ignored when IgnoreFields is true.
        builder.InsertField("QUOTE", "Hello field!");

        // Configure find/replace options to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("Hello", "Hi", options);

        // Verify that at least one replacement occurred (the one outside the field).
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement outside of fields.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
