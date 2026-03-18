using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class ModifyAllLists
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Iterate through every List in the document.
        foreach (List list in doc.Lists)
        {
            // Example uniform modification: restart numbering at each section.
            list.IsRestartAtEachSection = true;

            // Apply the same formatting to each level of the list.
            foreach (ListLevel level in list.ListLevels)
            {
                // Set the font color for the list label.
                level.Font.Color = Color.DarkBlue;

                // Ensure the list uses Arabic numbers for all levels.
                level.NumberStyle = NumberStyle.Arabic;
            }
        }

        // Update list labels so that changes are reflected in the document.
        doc.UpdateListLabels();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
