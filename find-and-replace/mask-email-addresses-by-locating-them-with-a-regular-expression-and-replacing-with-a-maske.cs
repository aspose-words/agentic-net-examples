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

        // Add sample paragraphs that contain e‑mail addresses.
        builder.Writeln("Contact us at support@example.com for assistance.");
        builder.Writeln("Send feedback to feedback@mydomain.org or sales@mydomain.org.");
        builder.Writeln("No email here, just text.");

        // Define a regular expression that matches typical e‑mail addresses.
        Regex emailRegex = new Regex(@"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", RegexOptions.Compiled);

        // Replace every e‑mail address with a masked placeholder.
        int replacementCount = doc.Range.Replace(emailRegex, "[masked email]");

        // Ensure that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("No e‑mail addresses were found to mask.");

        // Save the modified document to the local file system.
        const string outputPath = "MaskedEmails.docx";
        doc.Save(outputPath);
    }
}
