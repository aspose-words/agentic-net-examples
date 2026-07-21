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

        // Add a sample procedural module with some dummy code.
        VbaModule sampleModule = new VbaModule();
        sampleModule.Name = "SampleModule";
        sampleModule.Type = VbaModuleType.ProceduralModule;
        sampleModule.SourceCode = @"
Sub TestMacro()
    MsgBox ""Hello from VBA!""
End Sub
";
        doc.VbaProject.Modules.Add(sampleModule);

        // Add another module to demonstrate handling multiple modules.
        VbaModule anotherModule = new VbaModule();
        anotherModule.Name = "AnotherModule";
        anotherModule.Type = VbaModuleType.ProceduralModule;
        anotherModule.SourceCode = @"
Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
";
        doc.VbaProject.Modules.Add(anotherModule);

        // Define the standardized logging routine to be inserted.
        string loggingRoutine = @"
Sub LogError(errMsg As String)
    ' Simple logging routine – writes to the Immediate Window.
    Debug.Print ""Error: "" & errMsg
End Sub

";

        // Insert the logging routine at the beginning of each VBA module.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            // Guard against null source code.
            string originalSource = module.SourceCode ?? string.Empty;

            // If the logging routine already exists, skip insertion to avoid duplication.
            if (!originalSource.Contains("Sub LogError"))
            {
                module.SourceCode = loggingRoutine + originalSource;
            }
        }

        // Save the document in a macro‑enabled format.
        const string outputPath = "Output.docm";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to '{outputPath}' with logging routines added to VBA modules.");
    }
}
