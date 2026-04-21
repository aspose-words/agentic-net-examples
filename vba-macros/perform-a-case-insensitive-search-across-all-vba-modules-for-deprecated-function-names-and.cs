using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.docm");
        string updatedPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument_Updated.docm");

        // -----------------------------------------------------------------
        // 1. Create a new macro‑enabled document and add a VBA project.
        // -----------------------------------------------------------------
        Document doc = new Document();

        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // -----------------------------------------------------------------
        // 2. Add sample VBA modules that contain deprecated function names.
        // -----------------------------------------------------------------
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub Test()
    Dim result As Long
    result = OldFunc(10)   ' Deprecated function
    MsgBox result
End Sub

Function OldFunc(value As Long) As Long
    OldFunc = value * 2
End Function"
        };
        vbaProject.Modules.Add(module1);

        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub AnotherTest()
    Dim txt As String
    txt = oldfunc(5)   ' Different case, same deprecated name
    Debug.Print txt
End Sub"
        };
        vbaProject.Modules.Add(module2);

        // Save the initial document (optional, just to illustrate the workflow).
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 3. Perform a case‑insensitive search & replace across all modules.
        //    Replace the deprecated function name "OldFunc" with "NewFunc".
        // -----------------------------------------------------------------
        const string deprecatedName = "OldFunc";
        const string replacementName = "NewFunc";

        // Ensure the document actually has a VBA project before proceeding.
        if (doc.HasMacros && doc.VbaProject != null)
        {
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Replace all occurrences of the deprecated name, ignoring case.
                string updatedSource = Regex.Replace(
                    source,
                    $@"\b{Regex.Escape(deprecatedName)}\b",
                    replacementName,
                    RegexOptions.IgnoreCase);

                // Assign the modified source back to the module.
                module.SourceCode = updatedSource;
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(updatedPath);
    }
}
