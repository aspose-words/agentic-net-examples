using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Directory to store temporary documents.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "SampleProject";
            doc.VbaProject = project;
        }

        // Create a VBA module with deprecated API calls.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub TestMacro()
    ' Deprecated call
    Call OldFunction()
    MsgBox ""Done""
End Sub

Function OldFunction()
    OldFunction = 42
End Function
";

        // Add the module to the project.
        doc.VbaProject.Modules.Add(module);

        // Save the original document (macro-enabled).
        string originalPath = Path.Combine(outputDir, "Original.docm");
        doc.Save(originalPath);

        // Load the document back.
        Document loadedDoc = new Document(originalPath);

        // Verify that a VBA project exists.
        if (loadedDoc.VbaProject != null)
        {
            // Iterate through all VBA modules.
            foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = vbaModule.SourceCode ?? string.Empty;

                // Replace deprecated API calls with updated equivalents.
                // Example: replace "OldFunction" with "NewFunction".
                string updatedSource = source.Replace("OldFunction", "NewFunction");

                // Update the module's source code only if a change was made.
                if (!source.Equals(updatedSource, StringComparison.Ordinal))
                {
                    vbaModule.SourceCode = updatedSource;
                }
            }

            // Save the updated document.
            string updatedPath = Path.Combine(outputDir, "Updated.docm");
            loadedDoc.Save(updatedPath);
        }

        // Indicate completion.
        Console.WriteLine("VBA modules processed and saved to: " + outputDir);
    }
}
