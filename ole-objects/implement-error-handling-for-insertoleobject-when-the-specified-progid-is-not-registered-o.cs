using System;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        string progId = "Non.Existent.ProgId";
        InsertOleObject(progId);
    }

    private static void InsertOleObject(string progId)
    {
        try
        {
            Type comType = Type.GetTypeFromProgID(progId);
            if (comType == null)
            {
                throw new COMException($"ProgID '{progId}' is not registered on this system.");
            }

            object comObject = Activator.CreateInstance(comType);
            Console.WriteLine($"Successfully created COM object for ProgID '{progId}'.");
            // Placeholder for actual OLE insertion logic.

            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch (COMException ex)
        {
            Console.WriteLine($"COM error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
