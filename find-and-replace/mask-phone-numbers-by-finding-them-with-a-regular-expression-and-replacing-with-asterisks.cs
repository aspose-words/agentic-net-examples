using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class PhoneNumberMasking
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing phone numbers.
        builder.Writeln("Contact list:");
        builder.Writeln("John Doe: 123-456-7890");
        builder.Writeln("Jane Smith: 987.654.3210");
        builder.Writeln("Bob Johnson: 555 123 4567");
        builder.Writeln("No phone here.");

        // Save the original document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just saved.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches common US phone number formats.
        Regex phoneRegex = new Regex(@"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b");

        // Replace each phone number with asterisks for privacy.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(phoneRegex, "*****", options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one phone number to be masked.");

        // Save the masked document.
        const string outputPath = "masked.docx";
        loaded.Save(outputPath);
    }
}
