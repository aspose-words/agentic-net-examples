using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the DOCX file that may contain VBA macros.
        string docPath = @"C:\Docs\Sample.docx";

        // Load the document using the Document constructor (lifecycle rule).
        Document doc = new Document(docPath);

        // Check if the loaded document has macros via the Document.HasMacros property.
        bool hasMacrosInDoc = doc.HasMacros;
        Console.WriteLine($"Document.HasMacros: {hasMacrosInDoc}");

        // Alternatively, detect macros without fully loading the document using FileFormatUtil.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(docPath);
        bool hasMacrosInInfo = formatInfo.HasMacros;
        Console.WriteLine($"FileFormatInfo.HasMacros: {hasMacrosInInfo}");

        // If needed, you can remove macros from the document.
        // Uncomment the following lines to strip macros and save the cleaned file.
        /*
        if (hasMacrosInDoc)
        {
            doc.RemoveMacros(); // Removes the VBA project.
            string cleanedPath = @"C:\Docs\Sample_NoMacros.docx";
            doc.Save(cleanedPath); // Save using the Document.Save method (lifecycle rule).
            Console.WriteLine($"Macros removed and document saved to: {cleanedPath}");
        }
        */
    }
}
