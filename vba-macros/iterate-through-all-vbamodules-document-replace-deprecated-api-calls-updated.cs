using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaModuleUpdater
{
    static void Main()
    {
        // Load the DOCM document that contains VBA macros.
        Document doc = new Document("Input.docm");

        // Ensure the document actually has a VBA project.
        if (doc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain any VBA project.");
            return;
        }

        // Iterate through each VBA module in the project.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            // Example: replace a deprecated API call "OldFunction" with the new one "NewFunction".
            // Adjust the strings below to match the actual APIs you need to update.
            string oldApi = "OldFunction";
            string newApi = "NewFunction";

            // Perform the replacement on the module's source code.
            if (!string.IsNullOrEmpty(module.SourceCode) && module.SourceCode.Contains(oldApi))
            {
                module.SourceCode = module.SourceCode.Replace(oldApi, newApi);
                Console.WriteLine($"Updated module: {module.Name}");
            }
        }

        // Save the modified document.
        doc.Save("Output.docm");
    }
}
