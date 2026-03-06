using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ApplyDefaultTemplate
{
    static void Main()
    {
        // Path to the folder where input and output files are stored.
        // Replace with your actual directory.
        string dataDir = @"C:\MyProject\Docs\";

        // Create a new blank document.
        Document doc = new Document();

        // Enable automatic style updating so that the template's styles are applied
        // when the document is opened in Microsoft Word.
        doc.AutomaticallyUpdateStyles = true;

        // Create a SaveOptions object appropriate for the output file extension.
        // The overload that takes a file name determines the concrete SaveOptions type.
        SaveOptions options = SaveOptions.CreateSaveOptions("Result.docx");

        // Specify the DOCM template that will be used as the default template.
        // This template will be attached automatically because the document has no attached template.
        options.DefaultTemplate = dataDir + "Template.docm";

        // Save the document using the configured SaveOptions.
        // The document will now have the default template applied.
        doc.Save(dataDir + "Result.docx", options);
    }
}
