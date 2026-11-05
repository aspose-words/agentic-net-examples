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

        // Insert a checkbox form field with a name, default checked state, and custom size (points).
        string fieldName = "AcceptTerms";
        bool defaultChecked = true;
        int sizeInPoints = 15; // Custom size – InsertCheckBox expects an int.

        builder.InsertCheckBox(fieldName, defaultChecked, sizeInPoints);

        // Validate that the checkbox was added and its properties are correct.
        FormField checkBox = doc.Range.FormFields[fieldName];
        if (checkBox == null)
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");

        if (!checkBox.Checked)
            throw new InvalidOperationException("Checkbox default checked state is not set correctly.");

        // Save the document to disk.
        doc.Save("CheckboxFormField.docx");
    }
}
