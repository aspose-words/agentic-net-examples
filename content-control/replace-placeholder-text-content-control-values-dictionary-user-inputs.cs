using System;
using System.Collections.Generic;
using Aspose.Words;

class ReplaceContentControlPlaceholders
{
    static void Main()
    {
        // Create a new document with placeholder text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear {{FullName}},");
        builder.Writeln("Address: {{Address}}");
        builder.Writeln("Date: {{Date}}");

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

            // Replace all occurrences of the placeholder.
            doc.Range.Replace(placeholder, replacement);
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
