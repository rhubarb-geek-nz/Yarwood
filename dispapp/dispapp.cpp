/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include <objbase.h>
#include <stdio.h>
#include <displib_h.h>

int main(int argc, char** argv)
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	if (SUCCEEDED(hr))
	{
		DIClientSecurity* clientSecurity = NULL;

		hr = CoCreateInstance(CLSID_CClientSecurity, NULL, CLSCTX_INPROC_SERVER, IID_DIClientSecurity, (void**)&clientSecurity);

		if (SUCCEEDED(hr))
		{
			CLSID clsid;

			hr = CLSIDFromProgID(L"RhubarbGeekNz.RunningMan", &clsid);

			if (SUCCEEDED(hr))
			{
				IUnknown* punk = NULL;

				hr = GetActiveObject(clsid, NULL, &punk);

				if (SUCCEEDED(hr))
				{
					IDispatch* dispatch = NULL;

					hr = punk->QueryInterface(IID_IDispatch, (void**)&dispatch);

					if (SUCCEEDED(hr))
					{
						VARIANT serverPrincName, authInfo;
						DWORD dwAuthnSvc = -1, dwAuthzSvc = -1, dwAuthnLevel = -1, dwImpLevel = -1, dwCapabilities = -1;

						VariantInit(&serverPrincName);
						VariantInit(&authInfo);

						hr = clientSecurity->QueryBlanket(dispatch, &dwAuthnSvc, &dwAuthzSvc, &serverPrincName, &dwAuthnLevel, &dwImpLevel, &authInfo, &dwCapabilities);

						if (SUCCEEDED(hr))
						{
							printf("%08lxd,%08lxd,%08lxd,%08lxd,%08lxd\n", dwAuthnSvc, dwAuthzSvc, dwAuthnLevel, dwImpLevel, dwCapabilities);

							if (serverPrincName.vt == VT_BSTR)
							{
								printf("%S\n", serverPrincName.bstrVal);
							}

							if (authInfo.vt == VT_PTR)
							{
								printf("%p\n", authInfo.byref);
							}
						}

						VariantClear(&serverPrincName);
					}

					if (SUCCEEDED(hr))
					{
						IDispatch* proxy = NULL;

						hr = clientSecurity->CopyProxy(dispatch, &proxy);

						if (SUCCEEDED(hr))
						{
							dispatch->Release();
							dispatch = proxy;
						}
					}

					if (SUCCEEDED(hr))
					{
						DISPID dispId;

						hr = clientSecurity->SetBlanket(dispatch, -1, 0, NULL, 4, 3, NULL, 0);

						if (SUCCEEDED(hr))
						{
							BSTR name = SysAllocString(L"GetMessage");
							hr = dispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId);
							SysFreeString(name);
						}

						for (int hint = 1; hint < 6; hint++)
						{
							if (SUCCEEDED(hr))
							{
								DISPPARAMS params = { 0,0,0,0 };
								VARIANT result;
								VariantInit(&result);
								UINT errArg;
								EXCEPINFO ex;
								VARIANT arg;

								arg.vt = VT_I4;
								arg.intVal = hint;

								ZeroMemory(&ex, sizeof(ex));
								params.rgvarg = &arg;
								params.cArgs = 1;

								hr = dispatch->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &ex, &errArg);

								if (SUCCEEDED(hr) && result.vt == VT_BSTR)
								{
									printf("%S\n", result.bstrVal);
								}

								VariantClear(&result);
							}
							else
							{
								break;
							}
						}

						dispatch->Release();
					}

					punk->Release();
				}
			}

			clientSecurity->Release();
		}

		CoUninitialize();
	}

	if (FAILED(hr))
	{
		fprintf(stderr, "0x%lx\n", (long)hr);
	}

	return FAILED(hr);
}
