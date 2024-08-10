# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param([switch]$UnregServer)

trap
{
	throw $PSItem
}

$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

$ProcessArchitecture = $Env:PROCESSOR_ARCHITECTURE.ToLower()

switch ($ProcessArchitecture)
{
	'amd64' { $ProcessArchitecture = 'x64' }
}

$ProgramFiles = $Env:ProgramFiles

$CompanyDir = Join-Path -Path $ProgramFiles -ChildPath 'rhubarb-geek-nz'
$ProductDir = Join-Path -Path $CompanyDir -ChildPath 'Yarwood'
$InstallDir = Join-Path -Path $ProductDir -ChildPath $ProcessArchitecture
$DllName = 'RhubarbGeekNzYarwood.dll'
$DllPath = Join-Path -Path $InstallDir -ChildPath $DllName

$CLSID = '{9B520C14-EA80-430A-A842-8AC36640FF09}'
$LIBID = '{D4316C17-27B9-4CC5-83BC-F6A9E089E452}'
$LIBVER = '0.0'
$IID = '{D1E59313-47BC-4C75-9C0D-26ECB2652A01}'
$PROGID = 'RhubarbGeekNz.ClientSecurity'

if ($UnregServer)
{
	$DllPath, $InstallDir, $ProductDir | ForEach-Object {
		$FilePath = $_
		if (Test-Path $FilePath)
		{
			Remove-Item -LiteralPath $FilePath
		}
	}

	if (Test-Path $CompanyDir)
	{
		$children = Get-ChildItem -LiteralPath $CompanyDir

		if (-not $children)
		{
			Remove-Item -LiteralPath $CompanyDir
		}
	}

	foreach ($RegistryPath in 
		"HKLM:\SOFTWARE\Classes\CLSID\$CLSID\InprocServer32",
		"HKLM:\SOFTWARE\Classes\CLSID\$CLSID",
		"HKLM:\SOFTWARE\Classes\$PROGID\CLSID",
		"HKLM:\SOFTWARE\Classes\$PROGID",
		"HKLM:\SOFTWARE\Classes\TypeLib\$LIBID\$LIBVER\0\win32",
		"HKLM:\SOFTWARE\Classes\TypeLib\$LIBID\$LIBVER\0\win64",
		"HKLM:\SOFTWARE\Classes\TypeLib\$LIBID\$LIBVER\0",
		"HKLM:\SOFTWARE\Classes\TypeLib\$LIBID\$LIBVER\FLAGS",
		"HKLM:\SOFTWARE\Classes\TypeLib\$LIBID\$LIBVER\HELPDIR",
		"HKLM:\SOFTWARE\Classes\TypeLib\$LIBID\$LIBVER",
		"HKLM:\SOFTWARE\Classes\TypeLib\$LIBID",
		"HKLM:\SOFTWARE\Classes\Interface\$IID\ProxyStubClsid32",
		"HKLM:\SOFTWARE\Classes\Interface\$IID\ProxyStubClsid",
		"HKLM:\SOFTWARE\Classes\Interface\$IID\TypeLib",
		"HKLM:\SOFTWARE\Classes\Interface\$IID"
	)
	{
		if (Test-Path $RegistryPath)
		{
			Remove-Item -Path $RegistryPath
		}
	}
}
else
{
	if (Test-Path $DllPath)
	{
		Write-Warning "$DllPath is already installed"
	}
	else
	{
		$SourceDir = Join-Path -Path $PSScriptRoot -ChildPath $ProcessArchitecture
		$SourceFile = Join-Path -Path $SourceDir -ChildPath $DllName

		if (-not (Test-Path $SourceFile))
		{
			Write-Error "$SourceFile does not exist"
		}
		else
		{
			$CompanyDir, $ProductDir, $InstallDir | ForEach-Object {
				$FilePath = $_
				if (-not (Test-Path $FilePath))
				{
					$Null = New-Item -Path $FilePath -ItemType 'Directory'
				}
			}

			Copy-Item $SourceFile $DllPath
		}

		$RegistryPath = "HKLM:\SOFTWARE\Classes\CLSID\$CLSID\InprocServer32"

		if (Test-Path $RegistryPath)
		{
			$null = Set-Item -Path $RegistryPath -Value $DllPath
		}
		else
		{
			$null = New-Item -Path $RegistryPath -Value $DllPath -Force
		}

		$null = New-ItemProperty -Path $RegistryPath -Name 'ThreadingModel' -Value 'Both' -PropertyType 'String'

		$RegistryPath = "HKLM:\SOFTWARE\Classes\$PROGID\CLSID"

		if (Test-Path $RegistryPath)
		{
			$null = Set-Item -Path $RegistryPath -Value $CLSID
		}
		else
		{
			$null = New-Item -Path $RegistryPath -Value $CLSID -Force
		}

		$RegistryPath = "HKLM:\SOFTWARE\Classes\Interface\$IID\ProxyStubClsid32"

		if (Test-Path $RegistryPath)
		{
			$null = Set-Item -Path $RegistryPath -Value '{00020424-0000-0000-C000-000000000046}'
		}
		else
		{
			$null = New-Item -Path $RegistryPath -Value '{00020424-0000-0000-C000-000000000046}' -Force
		}

		$RegistryPath = "HKLM:\SOFTWARE\Classes\Interface\$IID\TypeLib"

		if (Test-Path $RegistryPath)
		{
			$null = Set-Item -Path $RegistryPath -Value $LIBID
		}
		else
		{
			$null = New-Item -Path $RegistryPath -Value $LIBID -Force
		}

		$null = New-ItemProperty -Path $RegistryPath -Name 'Version' -Value $LIBVER -PropertyType 'String'

		Add-Type -TypeDefinition @"
			using System;
			using System.ComponentModel;
			using System.Runtime.InteropServices;

			namespace RhubarbGeekNz.Yarwood
			{
				public class InterOp
				{
					[DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
					private static extern void LoadTypeLibEx(string szFile, uint regkind, out IntPtr pptlib);

					public static void RegisterTypeLib(string path)
					{
						IntPtr punk;
						LoadTypeLibEx(path, 1, out punk);
						Marshal.Release(punk);
					}
				}
			}
"@

		[RhubarbGeekNz.Yarwood.InterOp]::RegisterTypeLib($DllPath)
	}
}
