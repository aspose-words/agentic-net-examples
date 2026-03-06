using System;
using Aspose.Words;
using Aspose.Words.Fields; // Added for FieldType enum

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SampleForm.docx";

        // Load the existing document. This uses the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Access a form field by its name (e.g., a checkbox named "CheckBox1").
        var checkBox = doc.Range.FormFields["CheckBox1"];
        if (checkBox != null && checkBox.Type == FieldType.FieldFormCheckBox)
        {
            // Set the checkbox to checked.
            checkBox.Checked = true;
        }

        // Access a text input form field named "TextField1".
        var textField = doc.Range.FormFields["TextField1"];
        if (textField != null && textField.Type == FieldType.FieldFormTextInput)
        {
            // Set the text value.
            textField.Result = "Aspose.Words";
        }

        // Save the modified document.
        string outputPath = @"C:\Docs\SampleForm_Modified.docx";
        doc.Save(outputPath);
    }
}
