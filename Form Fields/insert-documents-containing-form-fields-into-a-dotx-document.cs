using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOTX template that will receive the form fields.
        Document template = new Document("Template.dotx");

        // Paths to the source documents that already contain form fields.
        string[] sourceFiles = { "FormDoc1.docx", "FormDoc2.docx" };

        // DocumentBuilder is used to insert content into the template.
        DocumentBuilder builder = new DocumentBuilder(template);

        foreach (string filePath in sourceFiles)
        {
            // Load each source document.
            Document source = new Document(filePath);

            // Insert the entire source document at the current builder position.
            // Keep the original formatting of the source document.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);
        }

        // After insertion you can work with the combined collection of form fields.
        var allFormFields = template.Range.FormFields;
        Console.WriteLine($"Total form fields in the resulting document: {allFormFields.Count}");

        // Save the combined document as a DOTX template.
        template.Save("Result.dotx", SaveFormat.Dotx);
    }
}
