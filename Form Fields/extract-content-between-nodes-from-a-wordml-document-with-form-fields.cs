using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractFormFieldContent
{
    static void Main()
    {
        // Load the WORDML (or DOCX) document that contains form fields.
        // The path can be adjusted to point to the actual file location.
        string inputPath = @"C:\Docs\InputDocument.docx";
        Document doc = new Document(inputPath);

        // Iterate through all form fields in the document.
        // The FormFields collection is accessible via the document's Range.
        foreach (FormField formField in doc.Range.FormFields)
        {
            // The Result property contains the text that the form field displays.
            // For text input fields this is the user‑entered text,
            // for drop‑down fields it is the selected item, etc.
            string fieldName = formField.Name;
            string fieldResult = formField.Result;

            Console.WriteLine($"Form field \"{fieldName}\" contains: \"{fieldResult}\"");
        }

        // Optionally, save a copy of the document after processing.
        string outputPath = @"C:\Docs\ProcessedDocument.docx";
        doc.Save(outputPath);
    }
}
