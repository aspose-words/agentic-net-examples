using System;
using System.IO;
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

        // Add sample text containing email addresses, phone numbers, and URLs.
        builder.Writeln("Contact us at john.doe@example.com or jane_smith@domain.org.");
        builder.Writeln("Call us at 123-456-7890 or 555 123 4567.");
        builder.Writeln("Visit our website: https://www.example.com or http://test.org/page.");

        // Define regular expression patterns.
        Regex emailPattern = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.IgnoreCase);
        Regex phonePattern = new Regex(@"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b");
        Regex urlPattern   = new Regex(@"\bhttps?://[^\s]+\b", RegexOptions.IgnoreCase);

        // Perform sequential replacements.
        int emailReplacements = doc.Range.Replace(emailPattern, "[email]");
        int phoneReplacements = doc.Range.Replace(phonePattern, "[phone]");
        int urlReplacements   = doc.Range.Replace(urlPattern, "[url]");

        // Validate that each type of replacement occurred at least once.
        if (emailReplacements == 0)
            throw new InvalidOperationException("No email addresses were replaced.");
        if (phoneReplacements == 0)
            throw new InvalidOperationException("No phone numbers were replaced.");
        if (urlReplacements == 0)
            throw new InvalidOperationException("No URLs were replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedDocument.docx");
        doc.Save(outputPath);

        // Optional: write counts to console (no user interaction required).
        Console.WriteLine($"Email replacements: {emailReplacements}");
        Console.WriteLine($"Phone replacements: {phoneReplacements}");
        Console.WriteLine($"URL replacements: {urlReplacements}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
