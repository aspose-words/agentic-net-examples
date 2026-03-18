using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceContentControlPlaceholders
{
    static void Main()
    {
        // Load the template document that contains content controls with placeholder text.
        Document doc = new Document("Template.docx");

        // Dictionary of user inputs where the key is the placeholder name (without braces)
        // and the value is the text that should replace the placeholder.
        var userInputs = new Dictionary<string, string>
        {
            { "FullName", "John Doe" },
            { "Address", "123 Main Street" },
            { "Date", DateTime.Today.ToString("MMMM d, yyyy") }
        };

        // Iterate over each entry in the dictionary and replace the corresponding placeholder.
        // Placeholders are expected to be in the form {{PlaceholderName}} inside the document.
        foreach (var entry in userInputs)
        {
            string placeholder = $"{{{{{entry.Key}}}}}"; // e.g. {{FullName}}
            string replacement = entry.Value;

            // Use the Range.Replace method to replace all occurrences of the placeholder.
            // No special FindReplaceOptions are required for simple text replacement.
            doc.Range.Replace(placeholder, replacement);
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
