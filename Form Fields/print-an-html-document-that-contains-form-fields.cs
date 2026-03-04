using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Folder where the generated HTML will be saved.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a check box form field.
        // Parameters: name, isChecked, size (points).
        builder.InsertCheckBox("CheckBox", false, 15);

        // Insert a text input form field.
        // Parameters: name, type, default text, placeholder, max length.
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Enter text", 50);

        // Configure HTML save options to export form fields as interactive <input> elements.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportFormFields = true
        };

        // Save the document as HTML.
        string htmlPath = Path.Combine(artifactsDir, "FormFields.html");
        doc.Save(htmlPath, htmlOptions);

        // Read the generated HTML and output it to the console.
        string htmlContent = File.ReadAllText(htmlPath);
        Console.WriteLine("Generated HTML with interactive form fields:");
        Console.WriteLine(htmlContent);
    }
}
