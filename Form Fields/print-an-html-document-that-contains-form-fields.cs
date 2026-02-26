using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for TextFormFieldType

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        builder.Write("Check this box: ");
        builder.InsertCheckBox("MyCheckBox", false, 15);
        builder.Writeln();

        // Insert a text input form field.
        builder.Write("Enter name: ");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.Writeln();

        // Insert a combo box (drop‑down) form field.
        builder.Write("Choose fruit: ");
        builder.InsertComboBox("MyComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);
        builder.Writeln();

        // Set HTML save options to export form fields as interactive HTML elements.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportFormFields = true
        };

        // Define the output HTML file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormFields.html");

        // Save the document as HTML using the configured options.
        doc.Save(outputPath, htmlOptions);

        // Read the generated HTML and output it to the console.
        string htmlContent = File.ReadAllText(outputPath);
        Console.WriteLine(htmlContent);
    }
}
