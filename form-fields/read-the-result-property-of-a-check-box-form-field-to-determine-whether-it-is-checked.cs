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

        // Insert a checkbox form field named "MyCheckBox" and set it to checked.
        builder.Write("Check this box: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, 0);

        // Save the document (required by the rules).
        doc.Save("FormFields.docx");

        // Retrieve the checkbox form field from the collection by its name.
        FormField formField = doc.Range.FormFields["MyCheckBox"];
        if (formField == null)
            throw new InvalidOperationException("The expected form field was not found.");

        // For a checkbox, use the Checked property to determine its state.
        bool isChecked = formField.Checked;

        // Output the result.
        Console.WriteLine($"Checkbox '{formField.Name}' is {(isChecked ? "checked" : "unchecked")}.");
    }
}
