using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaImport
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a folder for temporary files.
            string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            Directory.CreateDirectory(dataDir);

            // Path to the external .bas file that will be imported.
            string basFilePath = Path.Combine(dataDir, "SampleModule.bas");

            // Create a simple VBA module file if it does not exist.
            if (!File.Exists(basFilePath))
            {
                string sampleVba = @"
Attribute VB_Name = ""SampleModule""
Sub HelloWorld()
    MsgBox ""Hello from imported VBA module!""
End Sub
";
                File.WriteAllText(basFilePath, sampleVba);
            }

            // Path where the macro‑enabled document will be saved.
            string docPath = Path.Combine(dataDir, "DocumentWithMacro.docm");

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
            VbaModule module = new VbaModule
            {
                Name = "ImportedModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = vbaSource
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(module);

            // Save the document in a macro‑enabled format.
            doc.Save(docPath);

            // Simple verification: reload the document and output module info.
            Document loadedDoc = new Document(docPath);
            VbaModule imported = loadedDoc.VbaProject.Modules["ImportedModule"];
            Console.WriteLine($"Module Name: {imported.Name}");
            Console.WriteLine("First 100 characters of source code:");
            Console.WriteLine(imported.SourceCode?.Substring(0, Math.Min(100, imported.SourceCode.Length)) ?? string.Empty);
        }
    }
}
