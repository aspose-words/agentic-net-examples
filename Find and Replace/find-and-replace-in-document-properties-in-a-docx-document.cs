using System;
using Aspose.Words;
using Aspose.Words.Properties;

class ReplaceDocumentProperties
{
    static void Main()
    {
        // Input and output file paths.
        string inputPath = @"Input.docx";
        string outputPath = @"Output.docx";

        // Text to find and its replacement.
        string pattern = "OldCompany";
        string replacement = "NewCompany";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // ----- Built‑in document properties -----
        foreach (DocumentProperty prop in doc.BuiltInDocumentProperties)
        {
            // Only process string values.
            if (prop.Value is string text && text.Contains(pattern, StringComparison.OrdinalIgnoreCase))
            {
                // Replace the pattern and assign the new value back.
                prop.Value = text.Replace(pattern, replacement, StringComparison.OrdinalIgnoreCase);
            }
        }

        // ----- Custom document properties -----
        foreach (DocumentProperty prop in doc.CustomDocumentProperties)
        {
            if (prop.Value is string text && text.Contains(pattern, StringComparison.OrdinalIgnoreCase))
            {
                prop.Value = text.Replace(pattern, replacement, StringComparison.OrdinalIgnoreCase);
            }
        }

        // Refresh all DOCPROPERTY fields so they display the updated values.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save(outputPath);
    }
}
