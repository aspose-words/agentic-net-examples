using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Path for the document that will contain the checkbox form field.
        const string filePath = "FormFields_CheckBox.docx";

        // -------------------------------------------------
        // Create a new document and insert a checkbox field.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox with the name "MyCheckBox", unchecked by default, size 50 points.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
        // Save the document so that it can be loaded later.
        doc.Save(filePath);

        // -------------------------------------------------
        // Load the document and read the checkbox state.
        // -------------------------------------------------
        Document loadedDoc = new Document(filePath);
        FormFieldCollection formFields = loadedDoc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Retrieve the checkbox by its name.
        FormField retrievedCheckBox = formFields["MyCheckBox"];
        if (retrievedCheckBox == null)
            throw new InvalidOperationException("The expected checkbox form field was not found.");

        // Use the Checked property to determine the state.
        bool isChecked = retrievedCheckBox.Checked;

        // The Result property contains "1" for checked and "0" for unchecked.
        string resultValue = retrievedCheckBox.Result;

        // Output the results.
        Console.WriteLine($"Checkbox \"{retrievedCheckBox.Name}\" checked: {isChecked}");
        Console.WriteLine($"Result property value: \"{resultValue}\"");
    }
}
