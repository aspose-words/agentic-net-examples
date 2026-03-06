using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ApplyDefaultConditionalBlock
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable automatic style updating.
        // This tells Aspose.Words to apply style changes from the attached template
        // each time the document is opened in Microsoft Word.
        doc.AutomaticallyUpdateStyles = true;

        // The document does not have an attached template (AttachedTemplate is empty).
        // When saving, we will specify a default template to be used.
        // This is useful for documents that were created programmatically and need
        // to inherit styles from an existing .dotx/.dotm template.
        SaveOptions options = SaveOptions.CreateSaveOptions("Result.docx");
        options.DefaultTemplate = @"C:\Templates\Business brochure.dotx"; // path to your .dotx/.dotm template

        // Save the document using the SaveOptions that contain the default template.
        doc.Save(@"C:\Output\Result.docx", options);
    }
}
