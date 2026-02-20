using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Define the variable names and the values to replace them with.
        var variableReplacements = new Dictionary<string, string>
        {
            { "CustomerName", "John Doe" },
            { "OrderNumber", "12345" },
            { "Address", "123 Main St, Anytown" }
        };

        // Update the document's variable collection.
        foreach (var kvp in variableReplacements)
        {
            // If the variable already exists, assign the new value; otherwise add it.
            if (doc.Variables.Contains(kvp.Key))
                doc.Variables[kvp.Key] = kvp.Value;
            else
                doc.Variables.Add(kvp.Key, kvp.Value);
        }

        // Refresh all DOCVARIABLE fields so they display the updated values.
        doc.Range.UpdateFields();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
