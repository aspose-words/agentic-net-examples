using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;

class PhoneNumberRenderer
{
    static void Main()
    {
        // The phone number to be validated and rendered.
        string phoneNumber = "+1 (555) 123-4567";

        // Regular expression that defines the required phone number format.
        // Example pattern: optional '+' followed by country code, optional spaces,
        // optional parentheses around area code, and groups of digits separated by spaces or hyphens.
        string pattern = @"^\+?\d{1,3}\s?\(?\d{3}\)?[\s-]\d{3}[\s-]\d{4}$";

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Validate the phone number using Regex.IsMatch before rendering.
        if (Regex.IsMatch(phoneNumber, pattern))
        {
            // Phone number matches the required format – render it.
            builder.Writeln($"Phone: {phoneNumber}");
        }
        else
        {
            // Phone number does not match – render a placeholder or error message.
            builder.Writeln("Invalid phone number");
        }

        // Save the document to disk.
        doc.Save("PhoneNumberOutput.docx");
    }
}
