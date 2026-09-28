using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with paragraphs containing the word "old".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first paragraph with old value.");
        builder.Writeln("Second paragraph also contains old.");
        builder.Writeln("Third paragraph without the keyword.");

        // Save the initial document (optional, just for reference).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform find-and-replace.
        Document loaded = new Document(inputPath);

        // Define the text to find and the replacement string.
        // The replacement string uses the metacharacter "\r" to insert a paragraph break
        // after the replaced text, effectively adding a line break.
        const string findText = "old";
        const string replaceText = "new\r";

        // Perform the replacement.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(findText, replaceText, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
