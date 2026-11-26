using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with email addresses, phone numbers, and URLs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello team,");
        builder.Writeln("Please contact John Doe at john.doe@example.com or Jane Smith at jane.smith@domain.org.");
        builder.Writeln("You can also reach us by phone: 123-456-7890 or 987.654.3210.");
        builder.Writeln("Visit our website at https://www.example.com or http://legacy-site.net for details.");
        builder.Writeln("Thank you.");

        // Save the original document (optional, just for reference).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Define regex patterns for email, phone number, and URL.
        Regex emailRegex = new Regex(@"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b", RegexOptions.Compiled);
        Regex phoneRegex = new Regex(@"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b", RegexOptions.Compiled);
        Regex urlRegex = new Regex(@"\bhttps?://[^\s]+", RegexOptions.Compiled);

        // Chain replacements: emails → phones → URLs.
        int emailReplacements = doc.Range.Replace(emailRegex, "[email]", new FindReplaceOptions());
        if (emailReplacements == 0)
            throw new InvalidOperationException("Expected at least one email address replacement.");

        int phoneReplacements = doc.Range.Replace(phoneRegex, "[phone]", new FindReplaceOptions());
        if (phoneReplacements == 0)
            throw new InvalidOperationException("Expected at least one phone number replacement.");

        int urlReplacements = doc.Range.Replace(urlRegex, "[url]", new FindReplaceOptions());
        if (urlReplacements == 0)
            throw new InvalidOperationException("Expected at least one URL replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
