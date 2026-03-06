using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source document, the template to apply, and the output PDF.
        string inputPath = @"C:\Docs\Input.docx";
        string templatePath = @"C:\Docs\Business brochure.dotx";
        string outputPath = @"C:\Docs\Result.pdf";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Enable automatic style updating so that the default template can be applied
        // when the document is saved.
        doc.AutomaticallyUpdateStyles = true;

        // Create a save options object appropriate for the output file extension.
        // This will return a PdfSaveOptions instance.
        SaveOptions options = SaveOptions.CreateSaveOptions(outputPath);

        // Set the default template that will be used if the document has no attached template.
        options.DefaultTemplate = templatePath;

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, options);
    }
}
