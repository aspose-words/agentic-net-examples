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
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Add a sample procedural module.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = "Sub TestMacro()\n    MsgBox \"Hello from Module1\"\nEnd Sub";
        doc.VbaProject.Modules.Add(module1);

        // Add a sample class module.
        VbaModule module2 = new VbaModule();
        module2.Name = "Class1";
        module2.Type = VbaModuleType.ClassModule;
        module2.SourceCode = "Public Sub ClassMethod()\n    MsgBox \"Hello from Class1\"\nEnd Sub";
        doc.VbaProject.Modules.Add(module2);

        // Define the standardized logging routine to prepend.
        string loggingRoutine = 
            "Sub LogError(errMsg As String)\n" +
            "    ' Simple logging routine\n" +
            "    Debug.Print \"Error: \" & errMsg\n" +
            "End Sub\n\n";

        // Insert the logging routine at the beginning of each module.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            // Guard against null source code.
            string originalSource = module.SourceCode ?? string.Empty;
            module.SourceCode = loggingRoutine + originalSource;
        }

        // Save the document as a macro-enabled file.
        string outputPath = "Output.docm";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
