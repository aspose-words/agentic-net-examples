using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaMacroUpdater
{
    public class Program
    {
        // Paths for the sample documents.
        private const string OriginalDocPath = "Sample.docm";
        private const string UpdatedDocPath = "Sample_Updated.docm";

        public static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create a macro‑enabled document with a VBA project.
            // -----------------------------------------------------------------
            Document doc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject project = new VbaProject
            {
                Name = "SampleProject"
            };
            doc.VbaProject = project;

            // Add a procedural module containing a deprecated API call.
            VbaModule module1 = new VbaModule
            {
                Name = "Module1",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub Test()
    ' Deprecated call
    Call OldFunction()
    MsgBox ""Hello from Module1""
End Sub"
            };
            doc.VbaProject.Modules.Add(module1);

            // Add another module with a different deprecated call.
            VbaModule module2 = new VbaModule
            {
                Name = "Module2",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub AnotherTest()
    ' Another deprecated call
    Dim result As String
    result = OldFunction()
    MsgBox result
End Sub"
            };
            doc.VbaProject.Modules.Add(module2);

            // Save the document in a macro‑enabled format.
            doc.Save(OriginalDocPath);

            // -----------------------------------------------------------------
            // Step 2: Load the document and replace deprecated API calls.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(OriginalDocPath);

            // Ensure the document actually contains a VBA project.
            if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
            {
                foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
                {
                    // Guard against null source code.
                    string source = vbaModule.SourceCode ?? string.Empty;

                    // Replace the deprecated call "OldFunction()" with "NewFunction()".
                    // You can add more replacement rules as needed.
                    if (source.Contains("OldFunction()"))
                    {
                        string updatedSource = source.Replace("OldFunction()", "NewFunction()");
                        vbaModule.SourceCode = updatedSource;
                    }
                }

                // Save the updated document.
                loadedDoc.Save(UpdatedDocPath);
            }

            // -----------------------------------------------------------------
            // Step 3: Simple verification output (non‑interactive).
            // -----------------------------------------------------------------
            // Load the updated document to confirm the changes.
            Document resultDoc = new Document(UpdatedDocPath);
            Console.WriteLine($"Document '{UpdatedDocPath}' has macros: {resultDoc.HasMacros}");
            foreach (VbaModule mod in resultDoc.VbaProject.Modules)
            {
                Console.WriteLine($"Module: {mod.Name}");
                Console.WriteLine(mod.SourceCode);
                Console.WriteLine(new string('-', 40));
            }
        }
    }
}
