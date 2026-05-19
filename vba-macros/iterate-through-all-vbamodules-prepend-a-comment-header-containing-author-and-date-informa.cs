using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Add a procedural module with some sample VBA code.
        VbaModule module1 = new VbaModule();
        module1.Name = "ModuleOne";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub";

        // Add a second module.
        VbaModule module2 = new VbaModule();
        module2.Name = "ModuleTwo";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = @"Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function";

        // Add modules to the VBA project.
        doc.VbaProject.Modules.Add(module1);
        doc.VbaProject.Modules.Add(module2);

        // Save the initial document (macro-enabled format).
        string originalPath = Path.Combine(outputDir, "Original.docm");
        doc.Save(originalPath, SaveFormat.Docm);

        // Iterate through all VBA modules and prepend a comment header.
        if (doc.HasMacros && doc.VbaProject != null)
        {
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Guard against null source code.
                string existingCode = module.SourceCode ?? string.Empty;

                // Build the comment header.
                string header = $"' Author: John Doe{Environment.NewLine}" +
                                $"' Date: {DateTime.Now:yyyy-MM-dd}{Environment.NewLine}{Environment.NewLine}";

                // Prepend the header to the existing source code.
                module.SourceCode = header + existingCode;
            }
        }

        // Save the modified document.
        string updatedPath = Path.Combine(outputDir, "Updated.docm");
        doc.Save(updatedPath, SaveFormat.Docm);
    }
}
