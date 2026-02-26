using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Path to the source text file containing the items.
        string txtPath = "input.txt";

        // ------------------------------
        // Create a document with numbered list items.
        // ------------------------------

        // Create a blank Word document.
        Document listDoc = new Document();

        // DocumentBuilder provides methods to insert content.
        DocumentBuilder listBuilder = new DocumentBuilder(listDoc);

        // Read each line from the text file.
        string[] lines = File.ReadAllLines(txtPath);

        foreach (string line in lines)
        {
            // Apply default numbering (1., 2., 3., …) to the current paragraph.
            listBuilder.ListFormat.ApplyNumberDefault();

            // Write the line and end the paragraph.
            listBuilder.Writeln(line);

            // Remove list formatting so the next line can start a new list entry.
            listBuilder.ListFormat.RemoveNumbers();
        }

        // Save the document that contains the list.
        listDoc.Save("OutputList.docx");

        // ------------------------------
        // Create a document with plain text (no list numbering).
        // ------------------------------

        Document plainDoc = new Document();
        DocumentBuilder plainBuilder = new DocumentBuilder(plainDoc);

        foreach (string line in lines)
        {
            // Write each line as a normal paragraph.
            plainBuilder.Writeln(line);
        }

        // Save the plain‑text document.
        plainDoc.Save("OutputPlain.docx");
    }
}
