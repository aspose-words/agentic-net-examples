using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class PhoneMaskCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        if (args?.Match?.Value == null)
            return ReplaceAction.Skip;

        // Create a mask of asterisks with the same length as the matched phone number.
        string mask = new string('*', args.Match.Value.Length);
        args.Replacement = mask;
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document containing phone numbers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact list:");
        builder.Writeln("John Doe: 123-456-7890");
        builder.Writeln("Jane Smith: 987 654 3210");
        builder.Writeln("Bob Johnson: 555.123.4567");
        builder.Writeln("No phone here.");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches common US phone number formats.
        Regex phoneRegex = new Regex(@"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b");

        // Set up the replace callback to mask each matched phone number.
        PhoneMaskCallback callback = new PhoneMaskCallback();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(phoneRegex, string.Empty, options);

        // Validate that at least one phone number was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one phone number to be masked.");

        // Save the masked document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
