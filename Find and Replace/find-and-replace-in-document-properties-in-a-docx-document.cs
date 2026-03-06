using System;
using Aspose.Words;
using Aspose.Words.Properties;

class ReplaceInDocumentProperties
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document(@"C:\Docs\InputDocument.docx");

        // Example: replace text in a built‑in property (Title).
        string oldTitle = doc.BuiltInDocumentProperties.Title;
        if (!string.IsNullOrEmpty(oldTitle))
        {
            // Replace all occurrences of "Draft" with "Final".
            string newTitle = oldTitle.Replace("Draft", "Final");
            doc.BuiltInDocumentProperties.Title = newTitle;
        }

        // Example: replace text in a custom document property named "ProjectName".
        const string customPropName = "ProjectName";
        if (doc.CustomDocumentProperties.Contains(customPropName))
        {
            // The value of a custom property is stored as an object; cast to string.
            object propValueObj = doc.CustomDocumentProperties[customPropName].Value;
            if (propValueObj is string propValueStr)
            {
                // Replace "Alpha" with "Beta" in the custom property value.
                string newPropValue = propValueStr.Replace("Alpha", "Beta");
                doc.CustomDocumentProperties[customPropName].Value = newPropValue;
            }
        }

        // Save the modified document.
        doc.Save(@"C:\Docs\OutputDocument.docx");
    }
}
