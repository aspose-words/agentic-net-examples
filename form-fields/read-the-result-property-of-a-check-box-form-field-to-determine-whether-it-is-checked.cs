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

        // Insert a checkbox form field with a known name.
        builder.Write("Check this box: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, 0);
        // Ensure the checkbox is saved with its current state.
        checkBox.Checked = true;

        // Save the document to a file (required by the rules).
        string filePath = Path.Combine(Environment.CurrentDirectory, "CheckBoxResult.docx");
        doc.Save(filePath);

        // Load the document back (simulating a separate read operation).
        Document loadedDoc = new Document(filePath);

        // Retrieve the checkbox form field by name.
        FormField loadedCheckBox = loadedDoc.Range.FormFields["MyCheckBox"];
        if (loadedCheckBox == null)
            throw new InvalidOperationException("The expected checkbox form field was not found.");

        // Read the Result property. For a checkbox, "1" means checked, "0" means unchecked.
        string result = loadedCheckBox.Result;
        bool isChecked = result == "1";

        // Output the determination.
        Console.WriteLine($"Checkbox '{loadedCheckBox.Name}' is {(isChecked ? "checked" : "unchecked")} (Result = \"{result}\").");
    }
}
