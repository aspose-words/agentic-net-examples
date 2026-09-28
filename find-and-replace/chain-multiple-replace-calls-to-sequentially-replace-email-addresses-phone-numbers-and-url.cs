using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing; // Needed for FindReplaceOptions
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing email addresses, phone numbers, and URLs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact us at support@example.com or sales@example.org.");
        builder.Writeln("Call us at 123-456-7890 or 987 654 3210.");
        builder.Writeln("Visit our website at https://www.example.com or http://example.org.");

        // Replace email addresses.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.Compiled);
        int emailReplaced = doc.Range.Replace(emailRegex, "[email redacted]", new FindReplaceOptions());
        if (emailReplaced == 0)
            throw new InvalidOperationException("No email addresses were replaced.");

        // Replace phone numbers.
        Regex phoneRegex = new Regex(@"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b", RegexOptions.Compiled);
        int phoneReplaced = doc.Range.Replace(phoneRegex, "[phone redacted]", new FindReplaceOptions());
        if (phoneReplaced == 0)
            throw new InvalidOperationException("No phone numbers were replaced.");

        // Replace URLs.
        Regex urlRegex = new Regex(@"\bhttps?://[^\s]+", RegexOptions.Compiled);
        int urlReplaced = doc.Range.Replace(urlRegex, "[url redacted]", new FindReplaceOptions());
        if (urlReplaced == 0)
            throw new InvalidOperationException("No URLs were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Write a JSON report of the replacement counts.
        var report = new
        {
            EmailsReplaced = emailReplaced,
            PhonesReplaced = phoneReplaced,
            UrlsReplaced = urlReplaced,
            OutputFile = outputPath
        };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("replacement-report.json", json);
    }
}
