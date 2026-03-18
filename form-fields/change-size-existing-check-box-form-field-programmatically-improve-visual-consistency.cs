using System;
using Aspose.Words;
using Aspose.Words.Fields;

class ChangeCheckBoxSize
{
    static void Main()
    {
        // Path to the source document that already contains a checkbox form field.
        const string inputPath = @"C:\Docs\Input.docx";

        // Path where the modified document will be saved.
        const string outputPath = @"C:\Docs\Output.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Retrieve the checkbox form field by its bookmark/name.
        // Replace "MyCheckBox" with the actual name of your checkbox.
        FormField checkBox = doc.Range.FormFields["MyCheckBox"];

        // Ensure the field exists and is a checkbox.
        if (checkBox != null && checkBox.Type == FieldType.FieldFormCheckBox)
        {
            // Enable explicit sizing for the checkbox.
            checkBox.IsCheckBoxExactSize = true;

            // Set the desired size in points (e.g., 30 points).
            checkBox.CheckBoxSize = 30.0;
        }
        else
        {
            Console.WriteLine("Checkbox form field not found or is not a checkbox.");
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
