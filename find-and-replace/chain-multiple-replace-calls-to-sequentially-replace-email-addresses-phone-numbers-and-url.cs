using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with email, phone number, and URL.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Contact us at john.doe@example.com or call 123-456-7890.");
        builder.Writeln("Visit https://example.com for more information.");
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath);

        // Load the document for processing.
        Document doc = new Document(inputPath);

        // Define regex patterns.
        const string emailPattern = @"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b";
        const string phonePattern = @"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b";
        const string urlPattern   = @"\bhttps?://[^\s]+";

        // Replace email addresses.
        int emailReplacements = doc.Range.Replace(
            new Regex(emailPattern),
            "[email redacted]",
            new FindReplaceOptions());

        if (emailReplacements == 0)
            throw new InvalidOperationException("Expected at least one email replacement.");

        // Replace phone numbers.
        int phoneReplacements = doc.Range.Replace(
            new Regex(phonePattern),
            "[phone redacted]",
            new FindReplaceOptions());

        if (phoneReplacements == 0)
            throw new InvalidOperationException("Expected at least one phone number replacement.");

        // Replace URLs.
        int urlReplacements = doc.Range.Replace(
            new Regex(urlPattern),
            "[url redacted]",
            new FindReplaceOptions());

        if (urlReplacements == 0)
            throw new InvalidOperationException("Expected at least one URL replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
