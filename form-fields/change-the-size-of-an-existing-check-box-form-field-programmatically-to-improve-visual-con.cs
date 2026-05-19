using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with an initial size.
        builder.Write("Sample checkbox: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 20);
        // Enable exact size handling so that we can set a custom size later.
        checkBox.IsCheckBoxExactSize = true;

        // Locate the checkbox by its name in the form fields collection.
        FormField field = doc.Range.FormFields["MyCheckBox"];
        if (field == null)
        {
            throw new InvalidOperationException("The checkbox form field 'MyCheckBox' was not found.");
        }

        // Ensure the field is treated as having an exact size.
        field.IsCheckBoxExactSize = true;
        // Change the size of the checkbox to improve visual consistency.
        field.CheckBoxSize = 30.0; // size in points

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ModifiedCheckbox.docx");
        doc.Save(outputPath);
    }
}
