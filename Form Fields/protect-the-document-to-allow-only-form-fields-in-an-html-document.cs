using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for TextFormFieldType enum

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert some form fields using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill out the form below:");
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name", 0);
        builder.InsertCheckBox("AgreeCheck", false, 15);
        builder.InsertComboBox("CountryCombo", new[] { "USA", "Canada", "UK" }, 0);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Configure HTML save options to export form fields as interactive <input> elements.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportFormFields = true
        };

        // Save the protected document as an HTML file.
        doc.Save("ProtectedFormFields.html", htmlOptions);
    }
}
