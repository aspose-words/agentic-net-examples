using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class ChangeCheckBoxSize
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with a known name.
        const string checkBoxName = "MyCheckBox";
        builder.InsertCheckBox(checkBoxName, false, true, 0);

        // Retrieve the checkbox form field by its name.
        FormField checkBox = doc.Range.FormFields[checkBoxName];

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

        // Define output path in the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");

        // Save the modified document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
