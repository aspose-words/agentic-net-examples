using System;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Example with a ProgID that is unlikely to be registered.
        string progId = "Non.Existent.ProgId";
        InsertOleObject(progId);
    }

    private static void InsertOleObject(string progId)
    {
        try
        {
            // Attempt to get the COM type for the specified ProgID.
            Type comType = Type.GetTypeFromProgID(progId, throwOnError: true);
            // Try to create an instance of the COM object.
            object comObject = Activator.CreateInstance(comType);
            Console.WriteLine($"Successfully created OLE object for ProgID '{progId}'.");
            // Release the COM object if applicable.
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch (COMException comEx)
        {
            Console.WriteLine($"COMException: ProgID '{progId}' is not registered. Details: {comEx.Message}");
        }
        catch (ArgumentException argEx)
        {
            Console.WriteLine($"ArgumentException: ProgID '{progId}' is invalid. Details: {argEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error for ProgID '{progId}': {ex.Message}");
        }
    }
}
