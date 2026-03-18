using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DefaultTemplateExample
{
    static void Main()
    {
        // Load or create a new document.
        Document doc = new Document();

        // Enable automatic style updating.
        doc.AutomaticallyUpdateStyles = true;

        // Example condition checks (replace with real logic).
        bool hasSpecificTemplate = false;   // placeholder condition
        bool hasAlternativeTemplate = false; // placeholder condition

        // Path placeholders – replace with actual file system paths.
        string specificTemplatePath = @"C:\Templates\SpecificTemplate.dotx";   // {TemplatePath}
        string alternativeTemplatePath = @"C:\Templates\AlternativeTemplate.dotx"; // {TemplatePath}
        string defaultTemplatePath = @"C:\Templates\DefaultTemplate.dotx";   // {DefaultTemplate}
        string outputPath = @"C:\Output\Result.docx";                       // {OutputPath}

        if (hasSpecificTemplate)
        {
            // Attach a specific template directly to the document.
            doc.AttachedTemplate = specificTemplatePath;
            // Save normally – the attached template will be used.
            doc.Save(outputPath);
        }
        else if (hasAlternativeTemplate)
        {
            // Attach an alternative template.
            doc.AttachedTemplate = alternativeTemplatePath;
            doc.Save(outputPath);
        }
        else
        {
            // No template attached – use SaveOptions.DefaultTemplate as a fallback.
            // Create a SaveOptions object appropriate for the .docx format.
            SaveOptions options = SaveOptions.CreateSaveOptions(outputPath);
            // Set the default template that will be applied during saving.
            options.DefaultTemplate = defaultTemplatePath;
            // Save the document with the default template applied.
            doc.Save(outputPath, options);
        }
    }
}
