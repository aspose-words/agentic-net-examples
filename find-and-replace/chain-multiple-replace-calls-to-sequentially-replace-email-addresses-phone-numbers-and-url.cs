using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample document with email addresses, phone numbers and URLs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact us at john.doe@example.com or jane_smith@domain.org.");
        builder.Writeln("Phone: +1-800-555-1234, 020 7946 0958.");
        builder.Writeln("Visit our website: https://www.example.com or http://sub.domain.co.uk/page.");

        // Save the document locally (create‑save workflow).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the saved document (load workflow).
        Document loaded = new Document(inputPath);

        // Replace email addresses.
        Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}");
        int emailReplacements = loaded.Range.Replace(emailRegex, "[email]", new FindReplaceOptions());
        if (emailReplacements == 0)
            throw new InvalidOperationException("Expected at least one email address replacement.");

        // Replace phone numbers.
        Regex phoneRegex = new Regex(@"\+?\d[\d\s\-\(\)]{5,}\d");
        int phoneReplacements = loaded.Range.Replace(phoneRegex, "[phone]", new FindReplaceOptions());
        if (phoneReplacements == 0)
            throw new InvalidOperationException("Expected at least one phone number replacement.");

        // Replace URLs.
        Regex urlRegex = new Regex(@"https?://[^\s]+");
        int urlReplacements = loaded.Range.Replace(urlRegex, "[url]", new FindReplaceOptions());
        if (urlReplacements == 0)
            throw new InvalidOperationException("Expected at least one URL replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Inform the user.
        Console.WriteLine($"Replacements performed: emails={emailReplacements}, phones={phoneReplacements}, urls={urlReplacements}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
