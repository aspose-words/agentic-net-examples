using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load a DOCM file that contains a VBA project.
        Document doc = new Document("Input.docm");

        // Access the collection of VBA references.
        VbaReferenceCollection references = doc.VbaProject.References;

        Console.WriteLine($"Total VBA references: {references.Count}");

        // List each reference with its type and extracted path.
        for (int i = 0; i < references.Count; i++)
        {
            VbaReference reference = references[i];
            string path = GetLibIdPath(reference);
            Console.WriteLine($"Reference {i}: Type = {reference.Type}, Path = {path}");
        }

        // Example: remove a reference that points to a broken DLL.
        const string brokenPath = @"C:\broken.dll";

        // Iterate backwards when removing items to avoid index issues.
        for (int i = references.Count - 1; i >= 0; i--)
        {
            VbaReference reference = references[i];
            if (GetLibIdPath(reference).Equals(brokenPath, StringComparison.OrdinalIgnoreCase))
            {
                references.RemoveAt(i);
                Console.WriteLine($"Removed reference at index {i}");
            }
        }

        // Save the modified document.
        doc.Save("Output.docm");
    }

    // Returns the file path part of a VbaReference's LibId, handling different reference types.
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

    // Extracts the path from a LibId string for Registered/Original/Control references.
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

    // Extracts the path from a LibId string for Project references.
    private static string GetLibIdProjectPath(string libIdProject)
    {
        return libIdProject != null ? libIdProject.Substring(3) : string.Empty;
    }
}
