using System;
using Aspose.Words;
using Aspose.Words.Math;

class RetrieveMathObjectTypes
{
    static void Main()
    {
        // Load an existing Word document that contains OfficeMath objects.
        // Replace the path with the actual location of your document.
        Document doc = new Document("OfficeMathSample.docx");

        // Get all OfficeMath nodes in the document (including those inside other nodes).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and determine its MathObjectType.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Retrieve the type of the OfficeMath object.
            MathObjectType type = officeMath.MathObjectType;

            // Determine whether the object is a fraction, radical, or something else.
            if (type == MathObjectType.Fraction)
            {
                Console.WriteLine("Found a Fraction.");
            }
            else if (type == MathObjectType.Radical)
            {
                Console.WriteLine("Found a Radical.");
            }
            else
            {
                Console.WriteLine($"Found a {type} object.");
            }
        }

        // Optionally, save the document if any modifications were made.
        // Here we simply save a copy to demonstrate the save rule.
        doc.Save("OfficeMathSample_Output.docx");
    }
}
