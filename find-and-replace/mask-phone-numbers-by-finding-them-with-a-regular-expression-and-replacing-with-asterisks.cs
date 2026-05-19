using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing phone numbers.
        builder.Writeln("Contact list:");
        builder.Writeln("John Doe: 123-456-7890");
        builder.Writeln("Jane Smith: (555) 123 4567");
        builder.Writeln("Bob Johnson: 987.654.3210");
        builder.Writeln("Alice Brown: 5551234567");

        // Save the original document (optional, for inspection).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(inputPath);

        // Define a regular expression that matches common US phone number formats.
        Regex phoneRegex = new Regex(@"\b(?:\d{3}[-.\s]?\d{3}[-.\s]?\d{4}|\(\d{3}\)\s?\d{3}[-.\s]?\d{4})\b");

        // Replace each phone number with a masked version (10 asterisks).
        int replacementCount = loadedDoc.Range.Replace(phoneRegex, "**********", new FindReplaceOptions());

        // Ensure that at least one phone number was masked.
        if (replacementCount == 0)
            throw new InvalidOperationException("No phone numbers were found to mask.");

        // Save the masked document.
        const string outputPath = "masked.docx";
        loadedDoc.Save(outputPath);
    }
}
