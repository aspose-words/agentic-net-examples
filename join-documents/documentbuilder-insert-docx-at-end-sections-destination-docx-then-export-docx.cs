using System;
using Aspose.Words;

class InsertDocAtSectionEnds
{
    static void Main()
    {
        // Create a destination document with a single section and a paragraph.
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("This is the destination document.");

        // Create a source document with a single section and a paragraph.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document.");

        // For each section in the destination document, append the content of every
        // section from the source document to the end of that destination section.
        foreach (Section destSection in dstDoc.Sections)
        {
            foreach (Section srcSection in srcDoc.Sections)
            {
                // AppendContent copies only the body content of the source section.
                destSection.AppendContent(srcSection);
            }
        }

        // Save the modified document.
        dstDoc.Save("Result.docx");
        Console.WriteLine("Result.docx has been created successfully.");
    }
}
