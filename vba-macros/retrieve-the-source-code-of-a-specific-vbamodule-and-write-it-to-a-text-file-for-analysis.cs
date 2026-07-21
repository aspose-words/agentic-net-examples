using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class RetrieveVbaModuleSource
{
    public static void Main()
    {
        // Define file names.
        string docPath = Path.Combine(Environment.CurrentDirectory, "Sample.docm");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ModuleSource.txt");
        string moduleName = "SampleModule";

        // -----------------------------------------------------------------
        // Create a new macro‑enabled document and add a VBA module if needed.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "AsposeProject";
            doc.VbaProject = project;
        }

        // Add a VBA module with some sample code if it does not already exist.
        if (doc.VbaProject.Modules[moduleName] == null)
        {
            VbaModule module = new VbaModule();
            module.Name = moduleName;
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";
            doc.VbaProject.Modules.Add(module);
        }

        // Save the document in macro‑enabled format.
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // Load the document (optional) and retrieve the source of the module.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        VbaProject vbaProject = loadedDoc.VbaProject;

        // Guard against missing project or module.
        string sourceCode = string.Empty;
        if (vbaProject != null)
        {
            VbaModule targetModule = vbaProject.Modules[moduleName];
            if (targetModule != null && !string.IsNullOrEmpty(targetModule.SourceCode))
                sourceCode = targetModule.SourceCode;
        }

        // Write the retrieved source code to a text file.
        File.WriteAllText(outputPath, sourceCode);
    }
}
