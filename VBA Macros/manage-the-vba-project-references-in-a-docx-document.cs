using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaReferenceManager
{
    // Path to the source document that contains a VBA project.
    private const string InputPath = @"C:\Docs\SourceWithVba.docm";

    // Path where the modified document will be saved.
    private const string OutputPath = @"C:\Docs\SourceWithVba_Modified.docm";

    // Path of the reference that should be removed from the VBA project.
    private const string BrokenReferencePath = @"X:\broken.dll";

    static void Main()
    {
        // Load the document (lifecycle rule: load).
        Document doc = new Document(InputPath);

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;

        // Get the collection of references.
        VbaReferenceCollection references = vbaProject.References;

        // Remove any reference whose LibId points to the broken path.
        for (int i = references.Count - 1; i >= 0; i--)
        {
            VbaReference reference = references[i];
            string path = GetLibIdPath(reference);

            if (string.Equals(path, BrokenReferencePath, StringComparison.OrdinalIgnoreCase))
                references.RemoveAt(i);
        }

        // Example: remove the first reference in the collection (if any).
        if (references.Count > 0)
            references.Remove(references[0]);

        // Save the modified document (lifecycle rule: save).
        doc.Save(OutputPath);
    }

    // Helper method to extract the file path from a VbaReference's LibId.
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

    // Extracts the path part from a LibId that represents an Automation type library.
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

    // Extracts the path part from a LibId that represents an external VBA project reference.
    private static string GetLibIdProjectPath(string libIdProject)
    {
        return libIdProject != null ? libIdProject.Substring(3) : string.Empty;
    }
}
