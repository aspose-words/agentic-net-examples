using System;
using Aspose.Words;
using Aspose.Words.Vba;

class ReplaceVbaProject
{
    static void Main()
    {
        // Paths to the documents.
        string inputPath = @"C:\Docs\Input.docm";      // Document whose VBA project will be replaced.
        string templatePath = @"C:\Docs\VbaTemplate.docm"; // Document containing the pre‑configured VBA project.
        string outputPath = @"C:\Docs\Output.docm";    // Resulting document.

        // Load the document that will receive the new VBA project.
        Document targetDoc = new Document(inputPath);

        // Load the template document that already contains the desired VBA project.
        Document templateDoc = new Document(templatePath);

        // Clone the VBA project from the template.
        VbaProject newVbaProject = templateDoc.VbaProject.Clone();

        // Replace the existing VBA project in the target document with the cloned one.
        targetDoc.VbaProject = newVbaProject;

        // Save the modified document.
        targetDoc.Save(outputPath);
    }
}
