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

        // Insert a check box form field named "MyCheckBox".
        // The second argument sets the initial checked state (true = checked).
        // The third argument specifies the size; 0 lets Word choose the size automatically.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, 0);

        // Save the document so that the form field persists.
        const string outputPath = "FormFields_CheckBox.docx";
        doc.Save(outputPath);

        // Retrieve the collection of form fields from the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Locate the check box by its name. Throw if it cannot be found.
        FormField field = formFields["MyCheckBox"];
        if (field == null)
            throw new InvalidOperationException("The expected check box form field was not found.");

        // Determine whether the check box is checked using the recommended Checked property.
        bool isChecked = field.Checked;

        // Output the result to the console.
        Console.WriteLine($"Check box \"{field.Name}\" is {(isChecked ? "checked" : "unchecked")}.");

        // (Optional) The Result property for a check box contains "1" for checked and "0" for unchecked.
        // string result = field.Result;
        // Console.WriteLine($"Result property value: {result}");
    }
}
