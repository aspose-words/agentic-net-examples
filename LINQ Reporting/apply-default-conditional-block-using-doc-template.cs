using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the template file.
        // Replace with an actual path on your machine.
        string myDir = @"C:\Templates\";

        // Path to the folder where the resulting document will be saved.
        // Replace with an actual path on your machine.
        string artifactsDir = @"C:\Output\";

        // Create a new blank document.
        Document doc = new Document();

        // Enable automatic style updating so that the document will pick up styles from the template.
        doc.AutomaticallyUpdateStyles = true;

        // The document has no attached template (AttachedTemplate is an empty string by default).

        // Create a SaveOptions object for the target file name.
        // The static factory method creates the appropriate SaveOptions subclass.
        SaveOptions options = SaveOptions.CreateSaveOptions("Document.DefaultTemplate.docx");

        // Specify the default template that will be applied when the document is saved.
        options.DefaultTemplate = myDir + "Business brochure.dotx";

        // Save the document using the configured options.
        doc.Save(artifactsDir + "Document.DefaultTemplate.docx", options);
    }
}
