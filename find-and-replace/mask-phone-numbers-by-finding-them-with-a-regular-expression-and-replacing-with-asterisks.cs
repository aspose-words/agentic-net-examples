using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class PhoneNumberMasking
{
    public static void Main()
    {
        // Create a sample document with phone numbers.
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);
        builder.Writeln("Customer list:");
        builder.Writeln("John Doe - 123-456-7890");
        builder.Writeln("Jane Smith - (555) 123 4567");
        builder.Writeln("Bob Johnson - 9876543210");
        builder.Writeln("Alice Brown - 555.987.6543");
        inputDoc.Save("input.docx");

        // Load the document for processing.
        Document doc = new Document("input.docx");

        // Regular expression to match common phone number formats.
        Regex phoneRegex = new Regex(@"\b(?:\+?\d{1,2}[\s-]?)?(?:\(?\d{3}\)?[\s.-]?)?\d{3}[\s.-]?\d{4}\b");

        // Replacement string – ten asterisks to mask the number.
        const string mask = "**********";

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(phoneRegex, mask, new FindReplaceOptions());

        // Ensure that at least one phone number was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No phone numbers were found to mask.");

        // Save the masked document.
        doc.Save("output.docx");

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Phone numbers masked: {replacedCount}");
    }
}
