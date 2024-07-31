/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace RhubarbGeekNzYarwood
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string progId = "RhubarbGeekNz.RunningMan";

#if NETFRAMEWORK
            object active = Marshal.GetActiveObject(progId);
#else
            Guid guid = Type.GetTypeFromProgID(progId).GUID;
            GetActiveObject(ref guid, IntPtr.Zero, out object active);
#endif

            new CClientSecurity().SetBlanket(active, uint.MaxValue, 0, null, 4, 3, null, 0);

            foreach (int hint in args.Length == 0 ? new int[] { 1, 2, 3, 4, 5 } : args.Select(t => Int32.Parse(t)).ToArray())
            {
                object result = active.GetType().InvokeMember("GetMessage", BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod, null, active, new object[] { hint });
                Console.WriteLine($"{hint} {result}");
            }
        }

#if NETFRAMEWORK
#else
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out Object ppunk);
#endif
    }
}
