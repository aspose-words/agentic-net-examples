using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Directory where the template and output files are located.
        string MyDir = @"C:\Templates\";
        string ArtifactsDir = @"C:\Output\";

        // Create a new blank document.
        Document doc = new Document();

        // Enable automatic style updating.
        doc.AutomaticallyUpdateStyles = true;

        // Verify that the document has no attached template.
        if (doc.AttachedTemplate != string.Empty)
            throw new InvalidOperationException("Document should not have an attached template.");

        // Create SaveOptions for the target file.
        SaveOptions options = SaveOptions.CreateSaveOptions("Document.DefaultTemplate.docx");

        // Specify the default template to be applied when saving.
        options.DefaultTemplate = MyDir + "Business brochure.dotx";

        // Save the document using the options that include the default template.
        doc.Save(ArtifactsDir + "Document.DefaultTemplate.docx", options);
    }
}
