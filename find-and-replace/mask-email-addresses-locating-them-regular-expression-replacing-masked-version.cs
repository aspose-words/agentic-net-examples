using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class EmailMaskingExample
{
    static void Main()
    {
        // Create a new document and add some sample text containing e‑mail addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact us at support@example.com or sales@my-company.org.");
        builder.Writeln("Another address: user.name+tag@sub.domain.co.uk");

        // Define a regular expression that matches typical e‑mail addresses.
        Regex emailRegex = new Regex(
            @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
            RegexOptions.Compiled);

        // Replace every e‑mail address with a masked placeholder.
        // The Replace method that accepts a Regex and a replacement string is used.
        doc.Range.Replace(emailRegex, "[email masked]");

        // Save the modified document.
        doc.Save("MaskedEmails.docx");
    }
}
