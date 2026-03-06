using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceDocumentVariables
{
    static void Main()
    {
        // Paths to the source and destination DOCX files.
        string inputPath = @"C:\Docs\Input.docx";
        string outputPath = @"C:\Docs\Output.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Iterate through all variables defined in the document.
        foreach (KeyValuePair<string, string> variable in doc.Variables)
        {
            // Build the placeholder format used in the document body, e.g., _VariableName_.
            string placeholder = $"_{variable.Key}_";

            // Replace every occurrence of the placeholder with the variable's value.
            // Using the simple Replace(string, string) overload performs a case‑insensitive, whole‑document replace.
            doc.Range.Replace(placeholder, variable.Value);
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
