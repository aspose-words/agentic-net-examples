using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field that the user can fill in.
        builder.Write("Please enter your name: ");
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 0);

        // Protect the whole document so that only form fields are editable.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Set HTML save options to export form fields as interactive <input> tags.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportFormFields = true
        };

        // Save the protected document as an HTML file.
        doc.Save("ProtectedFormFields.html", htmlOptions);
    }
}
