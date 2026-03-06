using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ApplyDefaultConditionalBlock
{
    static void Main()
    {
        // Path to the folder that contains input templates and output folder.
        string MyDir = @"C:\Docs\Input\";
        string ArtifactsDir = @"C:\Docs\Output\";

        // Load the WORDML template document.
        Document doc = new Document(MyDir + "Template.docx"); // replace with actual WORDML file name

        // Enable automatic style updating. This will cause the document to pick up style changes
        // from the attached template (or the default template if none is attached) when opened.
        doc.AutomaticallyUpdateStyles = true;

        // The document does not have an attached template (AttachedTemplate is empty by default).
        // Use SaveOptions to specify a default template that will be applied during saving.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions("Result.docx");
        saveOptions.DefaultTemplate = MyDir + "Business brochure.dotx";

        // Optionally, use WordML2003SaveOptions to control the raw WORDML output.
        // Here we enable pretty formatting for easier readability of the saved XML.
        WordML2003SaveOptions wordMlOptions = new WordML2003SaveOptions
        {
            PrettyFormat = true,
            DefaultTemplate = MyDir + "Business brochure.dotx"
        };

        // Save the document with the default template applied.
        // If you need the WORDML format, use the WordML2003SaveOptions instance.
        // Otherwise, use the generic SaveOptions instance for DOCX output.
        doc.Save(ArtifactsDir + "Result.docx", saveOptions);
        // For WORDML output uncomment the following line:
        // doc.Save(ArtifactsDir + "Result.xml", wordMlOptions);
    }
}
