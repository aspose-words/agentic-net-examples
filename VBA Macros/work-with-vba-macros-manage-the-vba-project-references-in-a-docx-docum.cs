using System;
using Aspose.Words;
using Aspose.Words.Vba;

class ManageVbaReferences
{
    static void Main()
    {
        // Load a macro-enabled document.
        Document doc = new Document("VBA project.docm");

        // Access the VBA project inside the document.
        VbaProject vbaProject = doc.VbaProject;

        // Get the collection of VBA references.
        VbaReferenceCollection references = vbaProject.References;

        Console.WriteLine($"Initial references count: {references.Count}");

        // Define the path of a reference that should be removed.
        const string brokenPath = @"X:\broken.dll";

        // Iterate backwards so that removal does not affect the loop index.
        for (int i = references.Count - 1; i >= 0; i--)
        {
            VbaReference reference = references[i];
            string path = GetLibIdPath(reference);

            if (path.Equals(brokenPath, StringComparison.OrdinalIgnoreCase))
            {
                references.RemoveAt(i);
                Console.WriteLine($"Removed reference to {brokenPath}");
            }
        }

        Console.WriteLine($"Final references count: {references.Count}");

        // Save the document with the updated VBA references.
        doc.Save("VBA project Modified.docm");
    }

    // Returns the file path part of a VbaReference's LibId.
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

    // Extracts the path from a LibId that represents a type library reference.
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

    // Extracts the path from a LibId that represents a project reference.
    private static string GetLibIdProjectPath(string libIdProject)
    {
        return libIdProject != null ? libIdProject.Substring(3) : string.Empty;
    }
}
