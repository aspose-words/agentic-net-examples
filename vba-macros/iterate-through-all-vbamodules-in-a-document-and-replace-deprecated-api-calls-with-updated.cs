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

        // Create a VBA module with some sample code that contains a deprecated API call.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub ExampleMacro()
    ' Deprecated API call
    Call OldFunction()
    MsgBox ""Done""
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the initial document (optional, just to illustrate the before state).
        doc.Save("Original.docm");

        // Iterate through all VBA modules and replace deprecated API calls.
        if (doc.HasMacros && doc.VbaProject != null)
        {
            foreach (VbaModule vbaModule in doc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = vbaModule.SourceCode ?? string.Empty;

                // Replace the deprecated call "OldFunction" with the updated "NewFunction".
                string updatedSource = source.Replace("OldFunction", "NewFunction");

                // Update the module only if a change was made.
                if (!string.Equals(source, updatedSource, StringComparison.Ordinal))
                {
                    vbaModule.SourceCode = updatedSource;
                }
            }
        }

        // Save the document after modifications.
        doc.Save("Updated.docm");
    }
}
