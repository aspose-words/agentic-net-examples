using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load a macro-enabled document (DOCM) that contains a VBA project.
        string inputPath = @"C:\Docs\VBA project.docm";
        Document doc = new Document(inputPath);

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;

        // Get the collection of VBA references.
        VbaReferenceCollection references = vbaProject.References;
        Console.WriteLine($"Initial reference count: {references.Count}");

        // Iterate through all references and print their LibId paths.
        for (int i = 0; i < references.Count; i++)
        {
            VbaReference reference = references[i];
            string path = GetLibIdPath(reference);
            Console.WriteLine($"Reference {i}: Type={reference.Type}, Path={path}");
        }

        // Example: remove a reference that points to a broken DLL.
        const string brokenPath = @"X:\broken.dll";
        for (int i = references.Count - 1; i >= 0; i--)
        {
            VbaReference reference = references[i];
            if (GetLibIdPath(reference).Equals(brokenPath, StringComparison.OrdinalIgnoreCase))
            {
                references.RemoveAt(i);
                Console.WriteLine($"Removed broken reference at index {i}");
            }
        }

        // Example: add a new reference (registered type library) if needed.
        // Note: Aspose.Words does not provide a direct way to create a VbaReference,
        // so this part is illustrative. In practice you would manipulate the existing
        // references or use external tools to add new ones.

        // Save the modified document.
        string outputPath = @"C:\Docs\VBA project Modified.docm";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }

    // Returns the path component of a VbaReference's LibId based on its type.
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

    // Extracts the file path from a reference LibId string.
    private static string GetLibIdReferencePath(string libIdReference)
    {
        if (!string.IsNullOrEmpty(libIdReference))
        {
            string[] parts = libIdReference.Split('#');
            if (parts.Length > 3)
                return parts[3];
        }
        return string.Empty;
    }

    // Extracts the file path from a project LibId string.
    private static string GetLibIdProjectPath(string libIdProject)
    {
        return !string.IsNullOrEmpty(libIdProject) ? libIdProject.Substring(3) : string.Empty;
    }
}
