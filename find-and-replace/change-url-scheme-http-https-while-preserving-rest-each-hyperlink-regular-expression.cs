using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;

class ChangeHttpToHttps
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Iterate over all hyperlink fields in the document.
        foreach (FieldHyperlink hyperlink in doc.Range.Fields.OfType<FieldHyperlink>())
        {
            // Skip local (bookmark) hyperlinks – they have no external address.
            if (string.IsNullOrEmpty(hyperlink.Address))
                continue;

            // Replace the "http://" scheme with "https://" using a regular expression.
            // The pattern matches the scheme at the start of the address, case‑insensitively.
            string updated = Regex.Replace(
                hyperlink.Address,
                @"^http://",
                "https://",
                RegexOptions.IgnoreCase);

            // Apply the updated URL back to the field.
            hyperlink.Address = updated;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
