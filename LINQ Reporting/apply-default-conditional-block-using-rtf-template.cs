using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source RTF template.
        string rtfTemplatePath = @"C:\Templates\SourceTemplate.rtf";

        // Path to the default template that will be applied when the document has no attached template.
        string defaultTemplatePath = @"C:\Templates\DefaultTemplate.dotx";

        // Path where the resulting RTF document will be saved.
        string outputPath = @"C:\Output\ResultDocument.rtf";

        // Load the RTF template into a Document object.
        Document doc = new Document(rtfTemplatePath);

        // Enable automatic style updating so that the default template can be applied.
        doc.AutomaticallyUpdateStyles = true;

        // Create RtfSaveOptions and set the DefaultTemplate property.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // When the document does not have an attached template, this template will be used.
            DefaultTemplate = defaultTemplatePath,

            // Optional: keep the default behavior for old readers.
            ExportImagesForOldReaders = true
        };

        // Save the document as RTF using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
