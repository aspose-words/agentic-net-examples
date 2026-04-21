using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project; create one if missing.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "SampleProject";
            doc.VbaProject = project;
        }

        // Add a sample VBA module if the project contains none.
        if (doc.VbaProject.Modules.Count == 0)
        {
            VbaModule module = new VbaModule();
            module.Name = "SampleModule";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = @"
Sub Example()
    ' Deprecated API call
    Application.ActiveDocument.SaveAs ""C:\temp\old.doc""
    Selection.TypeText ""Hello""
End Sub
";
            doc.VbaProject.Modules.Add(module);
        }

        // Map of deprecated API calls to their updated equivalents.
        var replacements = new Dictionary<string, string>
        {
            { "Application.ActiveDocument.SaveAs", "ActiveDocument.SaveAs2" },
            { "Selection.TypeText", "Selection.TypeParagraph" }
        };

        // Iterate through all VBA modules and replace deprecated calls.
        foreach (VbaModule vbaModule in doc.VbaProject.Modules)
        {
            // Guard against null source code.
            string source = vbaModule.SourceCode ?? string.Empty;

            foreach (KeyValuePair<string, string> pair in replacements)
            {
                if (source.Contains(pair.Key))
                {
                    source = source.Replace(pair.Key, pair.Value);
                }
            }

            vbaModule.SourceCode = source;
        }

        // Save the modified document as a macro-enabled file.
        const string outputPath = "UpdatedDocument.docm";
        doc.Save(outputPath);
    }
}
