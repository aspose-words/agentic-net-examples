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
        project.Name = "Aspose.Project";
        doc.VbaProject = project;

        // Add a simple procedural module.
        VbaModule module = new VbaModule();
        module.Name = "TestModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub";
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        string filePath = "MacroWithReference.docm";
        doc.Save(filePath);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(filePath);

        // Verify the presence of a reference to the Microsoft Excel Object Library.
        bool hasExcelReference = false;
        foreach (VbaReference reference in loadedDoc.VbaProject.References)
        {
            // The LibId of the Excel reference typically contains the word "EXCEL".
            if (reference.LibId != null && reference.LibId.IndexOf("EXCEL", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                hasExcelReference = true;
                break;
            }
        }

        // Output the verification result.
        Console.WriteLine("Excel reference present: " + hasExcelReference);
    }
}
