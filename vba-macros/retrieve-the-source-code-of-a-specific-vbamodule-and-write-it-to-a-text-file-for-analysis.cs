using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path for the macro‑enabled document.
        string docPath = Path.Combine(artifactsDir, "Sample.docm");

        // -----------------------------------------------------------------
        // Create a new document with a VBA project and a single module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Create a procedural VBA module with some simple code.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub"
        };
        project.Modules.Add(module);

        // Save the document in macro‑enabled format.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Load the document and retrieve the source code of the target module.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        VbaProject loadedProject = loadedDoc.VbaProject;

        if (loadedProject != null)
        {
            // Access the module by name; guard against null.
            VbaModule targetModule = loadedProject.Modules["SampleModule"];
            string sourceCode = targetModule?.SourceCode ?? string.Empty;

            // Write the source code to a text file for analysis.
            string txtPath = Path.Combine(artifactsDir, "SampleModuleSource.txt");
            File.WriteAllText(txtPath, sourceCode);
        }
    }
}
