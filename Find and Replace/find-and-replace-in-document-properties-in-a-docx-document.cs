using System;
using Aspose.Words;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        const string pattern = "_Company_";
        const string replacement = "Contoso Ltd.";

        // Replace occurrences in built‑in document properties.
        foreach (DocumentProperty prop in doc.BuiltInDocumentProperties)
        {
            if (prop.Value is string str && str.Contains(pattern))
                prop.Value = str.Replace(pattern, replacement);
        }

        // Replace occurrences in custom document properties.
        foreach (DocumentProperty prop in doc.CustomDocumentProperties)
        {
            if (prop.Value is string str && str.Contains(pattern))
                prop.Value = str.Replace(pattern, replacement);
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
