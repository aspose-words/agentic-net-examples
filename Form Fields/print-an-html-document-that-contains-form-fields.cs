using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        builder.Write("Check this box: ");
        builder.InsertCheckBox("CheckBox1", false, 15);
        builder.InsertParagraph();

        // Insert a combo box (drop‑down) form field.
        builder.Write("Choose an option: ");
        builder.InsertComboBox("ComboBox1", new[] { "Option1", "Option2", "Option3" }, 0);
        builder.InsertParagraph();

        // Insert a text input form field.
        builder.Write("Enter text: ");
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Enter here", 50);
        builder.InsertParagraph();

        // Configure save options to export form fields as interactive HTML elements.
        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
        {
            ExportFormFields = true,   // Export as <input>, <select>, etc.
            PrettyFormat = true        // Make the HTML output readable.
        };

        // Define the output HTML file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormFields.html");

        // Save the document as HTML using the configured options.
        doc.Save(outputPath, saveOptions);

        // Read the generated HTML and print it to the console.
        string htmlContent = File.ReadAllText(outputPath);
        Console.WriteLine(htmlContent);
    }
}
