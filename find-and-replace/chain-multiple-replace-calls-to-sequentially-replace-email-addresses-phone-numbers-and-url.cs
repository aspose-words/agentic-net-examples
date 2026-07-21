using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required by the package list (no System.Drawing usage)
using Newtonsoft.Json;        // Required by the package list (not used in this example)

public class Program
{
    public static void Main()
    {
        // Create a sample document with email addresses, phone numbers and URLs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact us at john.doe@example.com or jane_smith@domain.org.");
        builder.Writeln("Call us at +1-800-555-1234 or (555) 123-4567.");
        builder.Writeln("Visit our website at https://www.example.com or http://test.org/page.");
        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Replace email addresses.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w+\b", RegexOptions.IgnoreCase);
        int emailReplacements = loaded.Range.Replace(emailRegex, "[email]", new FindReplaceOptions());
        if (emailReplacements == 0)
            throw new InvalidOperationException("Expected at least one email address replacement.");

        // Replace phone numbers.
        Regex phoneRegex = new Regex(@"\b\+?\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{1,4}[-.\s]?\d{1,9}\b");
        int phoneReplacements = loaded.Range.Replace(phoneRegex, "[phone]", new FindReplaceOptions());
        if (phoneReplacements == 0)
            throw new InvalidOperationException("Expected at least one phone number replacement.");

        // Replace URLs.
        Regex urlRegex = new Regex(@"\bhttps?://[^\s]+", RegexOptions.IgnoreCase);
        int urlReplacements = loaded.Range.Replace(urlRegex, "[url]", new FindReplaceOptions());
        if (urlReplacements == 0)
            throw new InvalidOperationException("Expected at least one URL replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Simple verification output (optional).
        Console.WriteLine($"Email replacements: {emailReplacements}");
        Console.WriteLine($"Phone replacements: {phoneReplacements}");
        Console.WriteLine($"URL replacements: {urlReplacements}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
