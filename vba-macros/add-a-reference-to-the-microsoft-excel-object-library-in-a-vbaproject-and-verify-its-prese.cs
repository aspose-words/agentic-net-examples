using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the temporary macro‑enabled document.
        const string docPath = "ExcelReferenceDemo.docm";

        // -----------------------------------------------------------------
        // 1. Create a new blank document and a VBA project.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // If the document does not already contain a VBA project, create one.
        if (!doc.HasMacros)
        {
            VbaProject project = new VbaProject
            {
                Name = "DemoProject"
            };
            doc.VbaProject = project;
        }

        // Add a simple VBA module so that the document is truly macro‑enabled.
        VbaModule module = new VbaModule
        {
            Name = "DemoModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the saved document and look for a reference to the
        //    Microsoft Excel Object Library.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Ensure the document actually has a VBA project.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain a VBA project.");
            return;
        }

        // The VbaReferenceCollection does not expose an Add method, so we
        // cannot programmatically insert a new reference. Instead we verify
        // whether a reference to Excel is already present.
        bool excelReferenceFound = false;

        foreach (VbaReference reference in loadedDoc.VbaProject.References)
        {
            // The LibId string contains the library identifier.
            // For the Excel Object Library it typically includes "EXCEL".
            if (!string.IsNullOrEmpty(reference.LibId) &&
                reference.LibId.IndexOf("EXCEL", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                excelReferenceFound = true;
                break;
            }
        }

        if (excelReferenceFound)
        {
            Console.WriteLine("Microsoft Excel Object Library reference is present in the VBA project.");
        }
        else
        {
            Console.WriteLine("Microsoft Excel Object Library reference is NOT present in the VBA project.");
            Console.WriteLine("Adding references programmatically is not supported by the current Aspose.Words VBA API.");
        }
    }
}
