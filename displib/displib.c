/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include <displib_h.h>

static LONG globalUsage;
static HMODULE globalModule;
static OLECHAR globalModuleFileName[260];

typedef struct CClientSecurityData
{
	IUnknown IUnknown;
	DIClientSecurity DIClientSecurity;
	LONG lUsage;
	IUnknown* lpOuter;
	ITypeInfo* piTypeInfo;
	IUnknown* punkMarshal;
} CClientSecurityData;

#define GetBaseObjectPtr(x,y,z)     (x *)(((char *)(void *)z)-(size_t)(char *)(&(((x*)NULL)->y)))

static STDMETHODIMP CClientSecurity_IUnknown_QueryInterface(IUnknown* pThis, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_NOINTERFACE;
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, IUnknown, pThis);

	if (IsEqualIID(riid, &IID_IDispatch) || IsEqualIID(riid, &IID_DIClientSecurity))
	{
		IUnknown_AddRef(pData->lpOuter);

		*ppvObject = &(pData->DIClientSecurity);

		hr = S_OK;
	}
	else
	{
		if (IsEqualIID(riid, &IID_IUnknown))
		{
			InterlockedIncrement(&pData->lUsage);

			*ppvObject = &(pData->IUnknown);

			hr = S_OK;
		}
		else
		{
			if (pData->punkMarshal && IsEqualIID(riid, &IID_IMarshal))
			{
				hr = IUnknown_QueryInterface(pData->punkMarshal, riid, ppvObject);
			}
			else
			{
				*ppvObject = NULL;
			}
		}
	}

	return hr;
}

static STDMETHODIMP_(ULONG) CClientSecurity_IUnknown_AddRef(IUnknown* pThis)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, IUnknown, pThis);
	return InterlockedIncrement(&pData->lUsage);
}

static STDMETHODIMP_(ULONG) CClientSecurity_IUnknown_Release(IUnknown* pThis)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, IUnknown, pThis);
	LONG res = InterlockedDecrement(&pData->lUsage);

	if (!res)
	{
		if (pData->piTypeInfo) IUnknown_Release(pData->piTypeInfo);
		if (pData->punkMarshal) IUnknown_Release(pData->punkMarshal);
		LocalFree(pData);
		InterlockedDecrement(&globalUsage);
	}

	return res;
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_QueryInterface(DIClientSecurity* pThis, REFIID riid, void** ppvObject)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, DIClientSecurity, pThis);
	return IUnknown_QueryInterface(pData->lpOuter, riid, ppvObject);
}

static STDMETHODIMP_(ULONG) CClientSecurity_DIClientSecurity_AddRef(DIClientSecurity* pThis)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, DIClientSecurity, pThis);
	return IUnknown_AddRef(pData->lpOuter);
}

static STDMETHODIMP_(ULONG) CClientSecurity_DIClientSecurity_Release(DIClientSecurity* pThis)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, DIClientSecurity, pThis);
	return IUnknown_Release(pData->lpOuter);
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_GetTypeInfoCount(DIClientSecurity* pThis, UINT* pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_GetTypeInfo(DIClientSecurity* pThis, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, DIClientSecurity, pThis);
	if (iTInfo) return DISP_E_BADINDEX;
	ITypeInfo_AddRef(pData->piTypeInfo);
	*ppTInfo = pData->piTypeInfo;

	return S_OK;
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_GetIDsOfNames(DIClientSecurity* pThis, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, DIClientSecurity, pThis);

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispGetIDsOfNames(pData->piTypeInfo, rgszNames, cNames, rgDispId);
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_Invoke(DIClientSecurity* pThis, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	CClientSecurityData* pData = GetBaseObjectPtr(CClientSecurityData, DIClientSecurity, pThis);

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispInvoke(pThis, pData->piTypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_QueryBlanket(DIClientSecurity* pThis, IDispatch* pProxy, DWORD *pAuthnSvc, DWORD *pAuthzSvc, VARIANT *pServerPrincName, DWORD *pAuthnLevel, DWORD *pImpLevel, VARIANT* pAuthInfo, DWORD *pCapabilties)
{
	HRESULT hr = E_INVALIDARG;

	if (pProxy)
	{
		IClientSecurity* client = NULL;

		hr = IUnknown_QueryInterface(pProxy, &IID_IClientSecurity, &client);

		if (SUCCEEDED(hr))
		{
			__try
			{
				LPOLESTR name = NULL;
				COAUTHINFO* authInfo = NULL;

				if (pServerPrincName)
				{
					pServerPrincName->vt = VT_EMPTY;
				}

				hr = IClientSecurity_QueryBlanket(client,
					(IUnknown*)(void*)pProxy,
					pAuthnSvc,
					pAuthzSvc, 
					pServerPrincName  ? &name : NULL,
					pAuthnLevel,
					pImpLevel, 
					pAuthInfo ? &authInfo : NULL,
					pCapabilties);

				if (SUCCEEDED(hr))
				{
					if (pServerPrincName && name)
					{
						pServerPrincName->bstrVal = SysAllocString(name);
						pServerPrincName->vt = VT_BSTR;
					}

					if (pAuthInfo)
					{
						pAuthInfo->vt = VT_PTR;
						pAuthInfo->byref = authInfo;
					}
				}

				if (name)
				{
					CoTaskMemFree(name);
				}
			}
			__finally
			{
				IClientSecurity_Release(client);
			}
		}
	}

	return hr;
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_SetBlanket(DIClientSecurity* pThis, IDispatch* pProxy, DWORD dwAuthnSvc, DWORD dwAuthzSvc, VARIANT *pServerPrincName, DWORD dwAuthnLevel, DWORD dwImpLevel, VARIANT *pAuthInfo, DWORD dwCapabilties)
{
	HRESULT hr = E_INVALIDARG;

	if (pProxy)
	{
		IClientSecurity* client = NULL;

		hr = IUnknown_QueryInterface(pProxy, &IID_IClientSecurity, &client);

		if (SUCCEEDED(hr))
		{
			__try
			{
				BSTR name = NULL;
				COAUTHINFO* authInfo = NULL;

				if (pServerPrincName && pServerPrincName->vt == VT_BSTR)
				{
					name = pServerPrincName->bstrVal;
				}

				if (pAuthInfo && pAuthInfo->vt == VT_PTR)
				{
					authInfo = pAuthInfo->byref;
				}

				hr = IClientSecurity_SetBlanket(client, (IUnknown*)(void*)pProxy, dwAuthnSvc, dwAuthzSvc, name, dwAuthnLevel, dwImpLevel, authInfo, dwCapabilties);
			}
			__finally
			{
				IClientSecurity_Release(client);
			}
		}
	}

	return hr;
}

static STDMETHODIMP CClientSecurity_DIClientSecurity_CopyProxy(DIClientSecurity* pThis, IDispatch* pProxy, IDispatch** ppCopy)
{
	HRESULT hr = E_INVALIDARG;

	if (pProxy)
	{
		IClientSecurity* client = NULL;

		hr = IUnknown_QueryInterface(pProxy, &IID_IClientSecurity, &client);

		if (SUCCEEDED(hr))
		{
			__try
			{
				hr = IClientSecurity_CopyProxy(client, (IUnknown*)(void*)pProxy, (void*)ppCopy);
			}
			__finally
			{
				IClientSecurity_Release(client);
			}
		}
	}

	return hr;
}

static IUnknownVtbl CClientSecurity_IUnknownVtbl =
{
	CClientSecurity_IUnknown_QueryInterface,
	CClientSecurity_IUnknown_AddRef,
	CClientSecurity_IUnknown_Release
};

static DIClientSecurityVtbl CClientSecurity_DIClientSecurityVtbl =
{
	CClientSecurity_DIClientSecurity_QueryInterface,
	CClientSecurity_DIClientSecurity_AddRef,
	CClientSecurity_DIClientSecurity_Release,
	CClientSecurity_DIClientSecurity_GetTypeInfoCount,
	CClientSecurity_DIClientSecurity_GetTypeInfo,
	CClientSecurity_DIClientSecurity_GetIDsOfNames,
	CClientSecurity_DIClientSecurity_Invoke,
	CClientSecurity_DIClientSecurity_QueryBlanket,
	CClientSecurity_DIClientSecurity_SetBlanket,
	CClientSecurity_DIClientSecurity_CopyProxy
};

static STDMETHODIMP CClassObject_CClientSecurity_IClassFactory_QueryInterface(IClassFactory* pThis, REFIID riid, void** ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IClassFactory))
	{
		InterlockedIncrement(&globalUsage);

		*ppvObject = pThis;

		return S_OK;
	}

	return E_NOINTERFACE;
}

static STDMETHODIMP_(ULONG) CClassObject_CClientSecurity_IClassFactory_AddRef(IClassFactory* pThis)
{
	return InterlockedIncrement(&globalUsage);
}

static STDMETHODIMP_(ULONG) CClassObject_CClientSecurity_IClassFactory_Release(IClassFactory* pThis)
{
	return InterlockedDecrement(&globalUsage);
}

static STDMETHODIMP CClassObject_CClientSecurity_IClassFactory_CreateInstance(IClassFactory* pThis, LPVOID punk, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_FAIL;

	if (punk == NULL || IsEqualIID(riid, &IID_IUnknown))
	{
		ITypeLib* piTypeLib = NULL;

		hr = LoadTypeLibEx(globalModuleFileName, REGKIND_NONE, &piTypeLib);

		if (SUCCEEDED(hr))
		{
			CClientSecurityData* pData = LocalAlloc(LMEM_ZEROINIT, sizeof(*pData));

			if (pData)
			{
				IUnknown* p = &(pData->IUnknown);
				InterlockedIncrement(&globalUsage);

				pData->IUnknown.lpVtbl = &CClientSecurity_IUnknownVtbl;
				pData->DIClientSecurity.lpVtbl = &CClientSecurity_DIClientSecurityVtbl;

				pData->lUsage = 1;
				pData->lpOuter = punk ? punk : p;

				hr = ITypeLib_GetTypeInfoOfGuid(piTypeLib, &IID_DIClientSecurity, &pData->piTypeInfo);

				if (SUCCEEDED(hr))
				{
					if (punk)
					{
						hr = S_OK;

						*ppvObject = p;
					}
					else
					{
						hr = CoCreateFreeThreadedMarshaler(p, &pData->punkMarshal);

						if (SUCCEEDED(hr))
						{
							hr = IUnknown_QueryInterface(p, riid, ppvObject);
						}

						IUnknown_Release(p);
					}
				}
				else
				{
					IUnknown_Release(p);
				}
			}

			if (piTypeLib)
			{
				ITypeLib_Release(piTypeLib);
			}
		}
	}

	return hr;
}

static STDMETHODIMP CClassObject_CClientSecurity_IClassFactory_LockServer(IClassFactory* pThis, BOOL bLock)
{
	if (bLock)
	{
		InterlockedIncrement(&globalUsage);
	}
	else
	{
		InterlockedDecrement(&globalUsage);
	}

	return S_OK;
}

static IClassFactoryVtbl CClassObject_CClientSecurity_IClassFactoryVtbl =
{
	CClassObject_CClientSecurity_IClassFactory_QueryInterface,
	CClassObject_CClientSecurity_IClassFactory_AddRef,
	CClassObject_CClientSecurity_IClassFactory_Release,
	CClassObject_CClientSecurity_IClassFactory_CreateInstance,
	CClassObject_CClientSecurity_IClassFactory_LockServer
};

static struct CClassObject {
	IClassFactoryVtbl* lpVtbl;
} CClassObject_CClientSecurity = { &CClassObject_CClientSecurity_IClassFactoryVtbl };

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppvObject)
{
	if (IsEqualCLSID(rclsid, &CLSID_CClientSecurity))
	{
		return IUnknown_QueryInterface((IUnknown*)&CClassObject_CClientSecurity, riid, ppvObject);
	}

	return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow(void)
{
	return globalUsage ? S_FALSE : S_OK;
}

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		globalModule = hModule;
		DisableThreadLibraryCalls(hModule);

		if (!GetModuleFileNameW(globalModule, globalModuleFileName, sizeof(globalModuleFileName) / sizeof(globalModuleFileName[0])))
		{
			return FALSE;
		}

		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}
