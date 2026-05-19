using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with email addresses, phone numbers and URLs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact us at john.doe@example.com or jane_smith@domain.org.");
        builder.Writeln("Phone: 123-456-7890, 555 123 4567, 800.555.1234.");
        builder.Writeln("Visit our website: https://www.example.com or http://test.org/page.");
        // Save the source document (required by the lifecycle rule).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Replace email addresses.
        const string emailPattern = @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b";
        int emailReplaced = loaded.Range.Replace(new Regex(emailPattern), "[email]", new FindReplaceOptions());
        if (emailReplaced == 0)
            throw new InvalidOperationException("Expected at least one email address replacement.");

        // Replace phone numbers (simple US formats).
        const string phonePattern = @"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b";
        int phoneReplaced = loaded.Range.Replace(new Regex(phonePattern), "[phone]", new FindReplaceOptions());
        if (phoneReplaced == 0)
            throw new InvalidOperationException("Expected at least one phone number replacement.");

        // Replace URLs.
        const string urlPattern = @"\bhttps?://[^\s]+";
        int urlReplaced = loaded.Range.Replace(new Regex(urlPattern), "[url]", new FindReplaceOptions());
        if (urlReplaced == 0)
            throw new InvalidOperationException("Expected at least one URL replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Output replacement counts (optional, non‑interactive).
        Console.WriteLine($"Email replacements: {emailReplaced}");
        Console.WriteLine($"Phone replacements: {phoneReplaced}");
        Console.WriteLine($"URL replacements: {urlReplaced}");
    }
}
