using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "FormField.docx");
        string updatedPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFieldUpdated.docx");

        // -------------------------------------------------
        // Create a new document and insert a text input form field.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Write("Enter your name: ");
        // Insert a regular text input form field with name "UserName".
        builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "John Doe", 0);
        builder.Writeln(); // End the paragraph.

        // Save the original document.
        doc.Save(originalPath);

        // -------------------------------------------------
        // Load the document and update the form field's value.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Locate the form field by name within the document's range.
        FormField textInput = loadedDoc.Range.FormFields["UserName"];
        if (textInput != null)
        {
            // Update the displayed value of the text input form field.
            textInput.Result = "Jane Smith";
        }

        // Save the updated document.
        loadedDoc.Save(updatedPath);

        // Optional: output the new value to the console for verification.
        Console.WriteLine($"Form field \"UserName\" updated to: {textInput?.Result}");
    }
}
