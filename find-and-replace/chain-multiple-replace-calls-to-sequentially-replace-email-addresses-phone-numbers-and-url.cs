using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Required package as per task specification

public class FindAndReplaceChainExample
{
    public static void Main()
    {
        // Create a sample document with email, phone number, and URL placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact Information:");
        builder.Writeln("Email: john.doe@example.com");
        builder.Writeln("Phone: +1-555-123-4567");
        builder.Writeln("Website: https://www.example.com");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Define regular expressions for email, phone number, and URL patterns.
        Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", RegexOptions.IgnoreCase);
        Regex phoneRegex = new Regex(@"\+?\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{1,4}[-.\s]?\d{1,9}", RegexOptions.IgnoreCase);
        Regex urlRegex = new Regex(@"https?://[^\s]+", RegexOptions.IgnoreCase);

        // Perform sequential replacements.
        int emailReplacements = loaded.Range.Replace(emailRegex, "[email protected]");
        if (emailReplacements == 0)
            throw new InvalidOperationException("Expected at least one email address replacement.");

        int phoneReplacements = loaded.Range.Replace(phoneRegex, "[phone]");
        if (phoneReplacements == 0)
            throw new InvalidOperationException("Expected at least one phone number replacement.");

        int urlReplacements = loaded.Range.Replace(urlRegex, "[url]");
        if (urlReplacements == 0)
            throw new InvalidOperationException("Expected at least one URL replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Write a simple JSON report of the replacement counts.
        var report = new
        {
            EmailReplacements = emailReplacements,
            PhoneReplacements = phoneReplacements,
            UrlReplacements = urlReplacements,
            OutputFile = outputPath
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("replacement_report.json", jsonReport);
    }
}
