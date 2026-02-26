using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX template.
        Document doc = new Document("Template.docx");

        // Iterate through all document variables.
        foreach (KeyValuePair<string, string> variable in doc.Variables)
        {
            // Define the placeholder format used in the document.
            // Example: ${VariableName}
            string placeholder = "${" + variable.Key + "}";

            // Replace each placeholder with its corresponding variable value.
            doc.Range.Replace(placeholder, variable.Value);
        }

        // Save the updated document.
        doc.Save("Result.docx");
    }
}
