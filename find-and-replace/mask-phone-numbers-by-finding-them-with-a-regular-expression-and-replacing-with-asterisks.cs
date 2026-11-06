using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package, not used directly
using Newtonsoft.Json;        // Required package, not used directly

public class PhoneNumberMasking
{
    public static void Main()
    {
        // Create a sample document with phone numbers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact list:");
        builder.Writeln("John Doe: 123-456-7890");
        builder.Writeln("Jane Smith: (987) 654-3210");
        builder.Writeln("Office: 555.123.4567");
        builder.Writeln("Emergency: +1-800-555-0199");
        // Save the original document (optional, for inspection).
        doc.Save("input.docx");

        // Define a regular expression that matches common phone number formats.
        // This pattern matches:
        // - 123-456-7890
        // - (123) 456-7890
        // - 123.456.7890
        // - +1-800-555-0199
        Regex phoneRegex = new Regex(
            @"(\+?\d{1,2}[-.\s]?)?(\(?\d{3}\)?[-.\s]?)\d{3}[-.\s]?\d{4}",
            RegexOptions.Compiled);

        // Perform the replacement, masking the phone numbers with asterisks.
        // The replacement keeps the same length format using a generic mask.
        string mask = "***-***-****";
        int replacedCount = doc.Range.Replace(phoneRegex, mask, new FindReplaceOptions());

        // Validate that at least one phone number was masked.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No phone numbers were found to mask.");
        }

        // Save the masked document.
        doc.Save("masked.docx");

        // Output the result count (optional, not required for non‑interactive execution).
        Console.WriteLine($"Phone numbers masked: {replacedCount}");
    }
}
