using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaImportExample
{
    public class Program
    {
        public static void Main()
        {
            // Path for the temporary VBA module file.
            string basFilePath = Path.Combine(Directory.GetCurrentDirectory(), "SampleModule.bas");

            // Create a simple VBA module file if it does not exist.
            if (!File.Exists(basFilePath))
            {
                string sampleVba = @"Sub HelloWorld()
    MsgBox ""Hello from imported VBA!""
End Sub";
                File.WriteAllText(basFilePath, sampleVba);
            }

            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has a VBA project.
            if (doc.VbaProject == null)
            {
                VbaProject project = new VbaProject();
                project.Name = "ImportedProject";
                doc.VbaProject = project;
            }

            // Read the VBA source code from the .bas file.
            string vbaSource = File.ReadAllText(basFilePath) ?? string.Empty;

            // Create a new VBA module and set its properties.
            VbaModule module = new VbaModule();
            module.Name = "ImportedModule";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = vbaSource;

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(module);

            // Save the document as a macro‑enabled file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithImportedModule.docm");
            doc.Save(outputPath);
        }
    }
}
