using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains VBA macros.
        Document doc = new Document("Input.docx");

        // Get the collection of VBA project references.
        VbaReferenceCollection references = doc.VbaProject.References;

        // Define the path of a reference that should be removed (example).
        const string brokenPath = @"C:\broken.dll";

        // Iterate backwards to safely remove items from the collection.
        for (int i = references.Count - 1; i >= 0; i--)
        {
            VbaReference reference = references[i];
            string path = GetReferencePath(reference);

            // Remove the reference if its path matches the broken path.
            if (path.Equals(brokenPath, StringComparison.OrdinalIgnoreCase))
                references.RemoveAt(i);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Extracts the file path from a VbaReference's LibId based on its type.
    private static string GetReferencePath(VbaReference reference)
    {
        switch (reference.Type)
        {
            case VbaReferenceType.Registered:
            case VbaReferenceType.Original:
            case VbaReferenceType.Control:
                return ExtractPathFromLibId(reference.LibId);
            case VbaReferenceType.Project:
                return ExtractProjectPath(reference.LibId);
            default:
                throw new ArgumentOutOfRangeException();
        }
    }

    // Parses the LibId string for standard references to obtain the file path.
    private static string ExtractPathFromLibId(string libId)
    {
        if (string.IsNullOrEmpty(libId))
            return string.Empty;

        string[] parts = libId.Split('#');
        return parts.Length > 3 ? parts[3] : string.Empty;
    }

    // Parses the LibId string for project references to obtain the file path.
    private static string ExtractProjectPath(string libId)
    {
        return libId != null ? libId.Substring(3) : string.Empty;
    }
}
