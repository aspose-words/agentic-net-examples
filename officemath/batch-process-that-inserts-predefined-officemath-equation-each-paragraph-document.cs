using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the source document (lifecycle rule: load)
        string inputPath = "Input.docx";
        Document doc = new Document(inputPath);

        // Predefined EQ field arguments – a simple fraction 1/2.
        // This string is written after the field separator to define the equation.
        const string eqArguments = @"\f(1,2)";

        // Process every paragraph in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Optionally skip empty paragraphs.
            if (string.IsNullOrWhiteSpace(paragraph.GetText()))
                continue;

            // Insert an EQ field at the end of the current paragraph.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(paragraph);
            // Insert the field without immediate update (lifecycle rule: insert field).
            Field field = builder.InsertField(FieldType.FieldEquation, false);
            FieldEQ fieldEQ = (FieldEQ)field;

            // Write the equation arguments after the field separator.
            builder.MoveTo(fieldEQ.Separator);
            builder.Write(eqArguments);

            // Convert the EQ field to an OfficeMath object.
            OfficeMath officeMath = fieldEQ.AsOfficeMath();
            if (officeMath != null)
            {
                // Display the equation on its own line.
                officeMath.DisplayType = OfficeMathDisplayType.Display;

                // Insert the OfficeMath node before the field start.
                fieldEQ.Start.ParentNode.InsertBefore(officeMath, fieldEQ.Start);

                // Remove the original EQ field.
                fieldEQ.Remove();
            }
        }

        // Save the modified document (lifecycle rule: save)
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
