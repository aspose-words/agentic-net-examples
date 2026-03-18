using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load a macro-enabled document.
        Document doc = new Document("Input.docm");

        // Verify that the document actually contains VBA macros.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Access the VBA project inside the document.
        VbaProject vbaProject = doc.VbaProject;

        // Get the collection of references for the VBA project.
        VbaReferenceCollection references = vbaProject.References;

        // Look for a reference that points to the Microsoft Excel Object Library.
        bool excelReferenceFound = false;
        foreach (VbaReference reference in references)
        {
            // Extract the path part of the LibId according to the reference type.
            string path = GetLibIdPath(reference);

            // The Excel library typically contains the word "Excel" in its path.
            if (!string.IsNullOrEmpty(path) && path.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                excelReferenceFound = true;
                Console.WriteLine($"Excel reference found: {path}");
                break;
            }
        }

        // If the reference is not present, report it.
        if (!excelReferenceFound)
        {
            Console.WriteLine("Excel reference not found. Adding a new reference is not supported via Aspose.Words VBA API.");
        }

        // Save the document (even if unchanged) to demonstrate the lifecycle usage.
        doc.Save("Output.docm");
    }

    // Helper method to obtain the path component of a LibId based on its type.
    private static string GetLibIdPath(VbaReference reference)
    {
        switch (reference.Type)
        {
            case VbaReferenceType.Registered:
            case VbaReferenceType.Original:
            case VbaReferenceType.Control:
                return GetLibIdReferencePath(reference.LibId);
            case VbaReferenceType.Project:
                return GetLibIdProjectPath(reference.LibId);
            default:
                throw new ArgumentOutOfRangeException();
        }
    }

    // Extracts the file path from a LibId that represents an Automation type library.
    private static string GetLibIdReferencePath(string libIdReference)
    {
        if (libIdReference != null)
        {
            string[] parts = libIdReference.Split('#');
            if (parts.Length > 3)
                return parts[3];
        }
        return string.Empty;
    }

    // Extracts the file path from a LibId that represents a VBA project reference.
    private static string GetLibIdProjectPath(string libIdProject)
    {
        return libIdProject != null ? libIdProject.Substring(3) : string.Empty;
    }
}
