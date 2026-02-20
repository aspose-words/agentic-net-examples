using System;
using Aspose.Words;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string oldText = "OldCompany";
        string newText = "NewCompany";

        // Replace the text in built‑in document properties.
        foreach (DocumentProperty prop in doc.BuiltInDocumentProperties)
        {
            // Only string‑type properties can be processed with Replace.
            if (prop.Type == PropertyType.String && prop.Value != null)
            {
                string value = prop.Value.ToString();
                if (value.Contains(oldText))
                    prop.Value = value.Replace(oldText, newText);
            }
        }

        // Replace the text in custom document properties.
        foreach (DocumentProperty prop in doc.CustomDocumentProperties)
        {
            if (prop.Type == PropertyType.String && prop.Value != null)
            {
                string value = prop.Value.ToString();
                if (value.Contains(oldText))
                    prop.Value = value.Replace(oldText, newText);
            }
        }

        // Refresh any DOCPROPERTY fields that display the changed values.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
