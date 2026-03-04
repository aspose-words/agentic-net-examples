using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceDocumentVariables
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Example: add some variables to the document (if they are not already present).
        // In a real scenario the variables may already exist in the document.
        doc.Variables.Add("FullName", "John Doe");
        doc.Variables.Add("Date", DateTime.Today.ToShortDateString());

        // Iterate through all variables and replace placeholders in the document body.
        // Placeholders are expected to be in the format {{VariableName}}.
        foreach (KeyValuePair<string, string> entry in doc.Variables)
        {
            string placeholder = $"{{{{{entry.Key}}}}}"; // e.g., {{FullName}}
            string replacement = entry.Value ?? string.Empty;

            // Perform a case‑insensitive replace of the placeholder with the variable value.
            // No special options are required for this simple replacement.
            doc.Range.Replace(placeholder, replacement);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
