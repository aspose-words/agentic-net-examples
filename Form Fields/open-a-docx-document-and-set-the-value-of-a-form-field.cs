using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        const string inputPath = "input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve the form field by its bookmark/name.
        // Replace "MyTextInput" with the actual name of your text input field.
        FormField textField = doc.Range.FormFields["MyTextInput"];

        if (textField != null)
        {
            // Set a new value for the text input form field.
            textField.SetTextInputValue("New value");
        }
        else
        {
            Console.WriteLine("Form field not found.");
        }

        // Update all fields in the document to reflect the changes.
        doc.UpdateFields();

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
