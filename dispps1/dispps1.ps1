# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param(
	$ProgID = 'RhubarbGeekNz.RunningMan',
	$Method = 'GetMessage',
	$Hint = @(1, 2, 3, 4, 5)
)

$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

Add-Type -TypeDefinition @"
	using System;
	using System.Runtime.InteropServices;
	using System.Runtime.InteropServices.ComTypes;

	namespace RhubarbGeekNz.RunningMan
	{
		public class InterOp
		{
  			[DllImport("oleaut32.dll", PreserveSig = false)]
  			private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out Object ppunk);

			public static object GetActiveObject(string progId)
			{
				object active;
				Guid guid = Type.GetTypeFromProgID(progId).GUID;
				GetActiveObject(ref guid, IntPtr.Zero, out active);
				return active;
			}
		}
	}
"@

$helloWorld = [RhubarbGeekNz.RunningMan.InterOp]::GetActiveObject($ProgID)

$clientSecurity = New-Object -ComObject 'RhubarbGeekNz.ClientSecurity'

$clientSecurity.SetBlanket($helloWorld, [UInt32]::MaxValue, 0, $null, 4, 3, $null, 0)

foreach ($h in $hint)
{
	$result = $helloWorld.$Method($h)

	"$h $result"
}
