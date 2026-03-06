using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains tracked changes.
        Document doc = new Document("Revisions.docx");

        // Iterate over revision groups – each group represents a consecutive set of revisions
        // made by the same author and of the same type.
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            Console.WriteLine($"Author: {group.Author}");
            Console.WriteLine($"Revision Type: {group.RevisionType}");
            Console.WriteLine($"Revision Text: {group.Text}");
            Console.WriteLine();
        }

        // Iterate over individual revisions for more detailed information.
        for (int i = 0; i < doc.Revisions.Count; i++)
        {
            Revision rev = doc.Revisions[i];
            Console.WriteLine($"Revision #{i + 1}");
            Console.WriteLine($"  Author: {rev.Author}");
            Console.WriteLine($"  Type: {rev.RevisionType}");
            Console.WriteLine($"  Date: {rev.DateTime}");
            // ParentNode may be null for style definition changes.
            string text = rev.ParentNode != null ? rev.ParentNode.GetText().Trim() : "<no text>";
            Console.WriteLine($"  Text: {text}");
            Console.WriteLine();
        }

        // Save the document (optional – here we just create a copy after processing).
        doc.Save("Revisions_Processed.docx");
    }
}
