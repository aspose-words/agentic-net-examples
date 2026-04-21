using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // Create a new VBA module with simple macro code.
        VbaModule vbaModule = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Update the module's source code to include error handling.
        // Guard against null source code.
        string currentSource = vbaModule.SourceCode ?? string.Empty;

        // Add "On Error Resume Next" at the beginning if it's not already present.
        if (!currentSource.Contains("On Error Resume Next"))
        {
            // Use Windows line endings for VBA code.
            vbaModule.SourceCode = "On Error Resume Next\r\n" + currentSource;
        }

        // Save the document in a macro‑enabled format.
        doc.Save("UpdatedMacroDocument.docm");
    }
}
