using System;
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

        // Add sample text containing an email address and a phone number.
        builder.Writeln("Contact information:");
        builder.Writeln("Email: john.doe@example.com");
        builder.Writeln("Phone: 123-456-7890");

        // Configure find‑replace options (no special flags needed for regex usage).
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace the email address with a placeholder using a regular‑expression pattern.
        int emailReplacements = doc.Range.Replace(
            new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"),
            "[redacted email]",
            options);

        // Replace the phone number with a placeholder using a regular‑expression pattern.
        int phoneReplacements = doc.Range.Replace(
            new Regex(@"\d{3}-\d{3}-\d{4}"),
            "[redacted phone]",
            options);

        // Validate that replacements occurred.
        if (emailReplacements == 0 && phoneReplacements == 0)
            throw new InvalidOperationException("No replacements were made.");

        // Save the modified document to the local folder.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
