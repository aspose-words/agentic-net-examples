using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file (load rule)
        Document doc = new Document("Input.docx");

        // Retrieve and display all revisions in the document
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime}");
            Console.WriteLine($"Type: {rev.RevisionType}");
            Console.WriteLine($"Text: {rev.ParentNode.GetText().Trim()}");
            Console.WriteLine(new string('-', 40));
        }

        // Save the document (save rule) – unchanged, just to satisfy lifecycle requirement
        doc.Save("Processed.docx");
    }
}
