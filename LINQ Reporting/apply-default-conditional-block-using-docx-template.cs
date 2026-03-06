using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the template file.
        // Replace with your actual directory.
        string myDir = @"C:\Templates\";

        // Path to the folder where the resulting document will be saved.
        // Replace with your actual directory.
        string artifactsDir = @"C:\Output\";

        // Create a new blank document.
        Document doc = new Document();

        // Enable automatic style updating so that styles are taken from the template.
        doc.AutomaticallyUpdateStyles = true;

        // The document has no attached template (AttachedTemplate is empty by default).

        // Create SaveOptions and specify the default template to use when saving.
        SaveOptions options = SaveOptions.CreateSaveOptions("Document.DefaultTemplate.docx");
        options.DefaultTemplate = System.IO.Path.Combine(myDir, "Business brochure.dotx");

        // Save the document using the SaveOptions that apply the default template.
        doc.Save(System.IO.Path.Combine(artifactsDir, "Document.DefaultTemplate.docx"), options);
    }
}
