using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with a default checked state.
        builder.Write("Check this box: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, 0);

        // Save the document that now contains the form field.
        const string outputPath = "CheckBoxResult.docx";
        doc.Save(outputPath);

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Retrieve the checkbox by its name.
        FormField field = formFields["MyCheckBox"];
        if (field == null)
            throw new InvalidOperationException("The expected checkbox form field was not found.");

        // Read the Result property. For a checkbox it is "1" when checked, otherwise "0".
        string result = field.Result ?? string.Empty;
        bool isChecked = result == "1";

        // Output the determination.
        Console.WriteLine($"Checkbox is {(isChecked ? "checked" : "unchecked")} (Result = \"{result}\")");
    }
}
