using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            VbaProject vbaProject = new VbaProject();
            vbaProject.Name = "SampleProject";
            doc.VbaProject = vbaProject;
        }

        // Create a new procedural module with a simple macro.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Update the source code to include error handling.
        VbaModule targetModule = doc.VbaProject.Modules["SampleModule"];
        string source = targetModule.SourceCode ?? string.Empty;

        // Prepend "On Error Resume Next" if it's not already present.
        if (!source.Contains("On Error Resume Next"))
        {
            source = "On Error Resume Next\r\n" + source;
        }

        targetModule.SourceCode = source;

        // Save the document as a macro-enabled file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedVbaModule.docm");
        doc.Save(outputPath);
    }
}
