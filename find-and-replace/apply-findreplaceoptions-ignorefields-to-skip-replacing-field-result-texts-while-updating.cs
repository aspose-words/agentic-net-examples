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
        builder.Writeln("Hello world!");

        // Insert a field whose result contains the word "Hello".
        // Field code: QUOTE, field result: "Hello again!"
        builder.InsertField("QUOTE", "Hello again!");

        // Configure find‑replace to ignore text inside fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Replace "Hello" with "Greetings" outside of fields.
        int replaced = doc.Range.Replace("Hello", "Greetings", options);
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Verify that the field result was not altered.
        string docText = doc.GetText().Trim();
        if (!docText.Contains("Hello again!"))
            throw new InvalidOperationException("Field result was unexpectedly replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
