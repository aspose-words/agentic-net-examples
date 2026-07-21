using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Add a couple of sample VBA modules.
        AddSampleModule(vbaProject, "ModuleOne", "Sub HelloWorld()\n    MsgBox \"Hello from ModuleOne\"\nEnd Sub");
        AddSampleModule(vbaProject, "ModuleTwo", "Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function");

        // Define author and date information for the header.
        const string author = "John Doe";
        string date = DateTime.Now.ToString("yyyy-MM-dd");
        string header = $"' Author: {author}\n' Date: {date}\n";

        // Iterate through all VBA modules and prepend the header.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            // Ensure source code is not null.
            string originalSource = module.SourceCode ?? string.Empty;

            // Prepend the header.
            module.SourceCode = header + originalSource;
        }

        // Save the document in a macro-enabled format.
        doc.Save("Output.docm");
    }

    // Helper method to create and add a VBA module with given name and source code.
    private static void AddSampleModule(VbaProject project, string moduleName, string sourceCode)
    {
        VbaModule module = new VbaModule();
        module.Name = moduleName;
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = sourceCode;
        project.Modules.Add(module);
    }
}
