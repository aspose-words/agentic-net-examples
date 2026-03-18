using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the document that contains a Table of Contents.
        Document doc = new Document("Input.docx");

        // Update all fields so that the TOC result text is current.
        doc.UpdateFields();

        // Iterate through every field in the document.
        foreach (Field field in doc.Range.Fields)
        {
            // Identify TOC fields by their type.
            if (field.Type == FieldType.FieldTOC)
            {
                // Cast to FieldToc to access TOC‑specific members.
                FieldToc tocField = (FieldToc)field;

                // The Result property holds the displayed TOC entries as plain text.
                string tocResult = tocField.Result;

                // Split the result into individual lines (each line is a TOC entry).
                string[] entries = tocResult.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                Console.WriteLine("Extracted TOC entries:");
                foreach (string entry in entries)
                {
                    Console.WriteLine(entry);
                }
            }
        }

        // Save the (potentially unchanged) document.
        doc.Save("Output.docx");
    }
}
