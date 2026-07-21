using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define a temporary folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the macro‑enabled document we will create.
        string docPath = Path.Combine(outputDir, "DocumentWithMacros.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "AsposeDemoProject"
        };
        doc.VbaProject = vbaProject;

        // Add a simple VBA module so the document actually contains macros.
        VbaModule module = new VbaModule
        {
            Name = "DemoModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub ShowMessage()
    MsgBox ""Hello from VBA!""
End Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        doc.Save(docPath);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(docPath);

        // Verify that the document indeed has a VBA project.
        if (!loadedDoc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA project.");
            return;
        }

        // Access the VBA project references collection.
        VbaReferenceCollection references = loadedDoc.VbaProject.References;

        // Check whether a reference to the Microsoft Excel Object Library is present.
        bool excelReferenceFound = false;
        foreach (VbaReference reference in references)
        {
            // The LibId of the Excel reference typically contains the word "EXCEL".
            if (!string.IsNullOrEmpty(reference.LibId) &&
                reference.LibId.IndexOf("EXCEL", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                excelReferenceFound = true;
                break;
            }
        }

        // Output the verification result.
        Console.WriteLine(excelReferenceFound
            ? "Microsoft Excel Object Library reference is present in the VBA project."
            : "Microsoft Excel Object Library reference is NOT present in the VBA project.");

        // Note: Adding a new reference programmatically is not supported directly by the Aspose.Words VBA API.
        // To include a reference, the VBA project must be prepared (e.g., in Microsoft Word) before loading.
    }
}
