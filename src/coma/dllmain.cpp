#include "stdafx.h"


#include "comaimpl.h"
#include "comaclsid.h"


HRESULT RegisterObject(HMODULE hModule, REFGUID guid, PCWSTR pszDescription, PCWSTR pszThreadingModel);

HRESULT UnregisterObject(const GUID& guid);

// Module Ref count
volatile ULONG  g_cRefModule = 0;

// Handle to the DLL's module
HMODULE g_hModule = NULL;

ULONG DllAddRef()
{
	return InterlockedIncrement(&g_cRefModule);
}

ULONG DllRelease()
{
	return InterlockedDecrement(&g_cRefModule);
}

class CClassFactory : public IClassFactory
{
public:
	static HRESULT CreateInstance(
		REFCLSID clsid,
		REFIID   riid,
		_COM_Outptr_ void **ppv
		)
	{
		*ppv = NULL;

		HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

		if (IsEqualGUID(clsid, CLSID_Coma))
		{
			IClassFactory *pClassFactory = new CClassFactory();
			if (!pClassFactory)
			{
				return E_OUTOFMEMORY;
			}
			hr = pClassFactory->QueryInterface(riid, ppv);
			pClassFactory->Release();
		}
		return hr;
	}

	//IUnknown methods
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppvObject)
	{
		HRESULT hr = S_OK;

		if (ppvObject == NULL)
		{
			return E_POINTER;
		}

		if (riid == IID_IUnknown)
		{
			*ppvObject = (IUnknown*)this;
		}
		else if (riid == IID_IClassFactory)
		{
			*ppvObject = (IClassFactory*)this;
		}
		else
		{
			*ppvObject = NULL;
			return E_NOINTERFACE;
		}

		AddRef();

		return hr;
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&m_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		long cRef = InterlockedDecrement(&m_cRef);
		if (cRef == 0)
		{
			delete this;
		}
		return cRef;
	}

	//IClassFactory Methods
	IFACEMETHODIMP CreateInstance(_In_ IUnknown *punkOuter, _In_ REFIID riid, _Outptr_ void **ppv)
	{
		return punkOuter ? CLASS_E_NOAGGREGATION : CComa::CreateInstance(riid, ppv);
	}

	IFACEMETHODIMP LockServer(BOOL fLock)
	{
		if (fLock)
		{
			DllAddRef();
		}
		else
		{
			DllRelease();
		}
		return S_OK;
	}

private:
	CClassFactory() :
		m_cRef(1)
	{
		DllAddRef();
	}

	~CClassFactory()
	{
		DllRelease();
	}

	long m_cRef;
};




STDAPI_(BOOL) DllMain(_In_opt_ HINSTANCE hinst, DWORD reason, _In_opt_ void *)
{
	if (reason == DLL_PROCESS_ATTACH)
	{
		g_hModule = (HMODULE)hinst;
		DisableThreadLibraryCalls(hinst);
	}
	return TRUE;
}

//
// DllCanUnloadNow
//
/////////////////////////////////////////////////////////////////////////
STDAPI DllCanUnloadNow()
{
	return (g_cRefModule == 0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////
//
// DllGetClassObject
//
/////////////////////////////////////////////////////////////////////////
_Check_return_
STDAPI  DllGetClassObject(
	_In_        REFCLSID rclsid,
	_In_        REFIID riid,
	_Outptr_    LPVOID FAR *ppv
	)
{
	return CClassFactory::CreateInstance(rclsid, riid, ppv);
}






STDMETHODIMP DllRegisterServer()
{


	// Register the CLSID for CoCreateInstance.
	HRESULT hr = RegisterObject(g_hModule, CLSID_Coma, TEXT("Coma"), TEXT("Both"));

	return hr;
}

STDMETHODIMP DllUnregisterServer()
{
	
	// Unregister the CLSID.
	UnregisterObject(CLSID_Coma);

	
	return S_OK;
}


// Converts a CLSID into a string with the form "CLSID\{clsid}"
STDMETHODIMP CreateObjectKeyName(REFGUID guid, _Out_writes_(cchMax) PWSTR pszName, DWORD cchMax)
{
	const DWORD chars_in_guid = 39;

	// convert CLSID uuid to string
	OLECHAR szCLSID[chars_in_guid];
	HRESULT hr = StringFromGUID2(guid, szCLSID, chars_in_guid);
	if (SUCCEEDED(hr))
	{
		// Create a string of the form "CLSID\{clsid}"
		hr = StringCchPrintf((STRSAFE_LPWSTR)pszName, cchMax, TEXT("Software\\Classes\\CLSID\\%ls"), szCLSID);
	}
	return hr;
}

// Creates a registry key (if needed) and sets the default value of the key
STDMETHODIMP CreateRegKeyAndValue(HKEY hKey, PCWSTR pszSubKeyName, PCWSTR pszValueName,
	PCWSTR pszData, PHKEY phkResult)
{
	*phkResult = NULL;
	LONG lRet = RegCreateKeyExW(
		hKey, pszSubKeyName,
		0, NULL, REG_OPTION_NON_VOLATILE,
		KEY_ALL_ACCESS, NULL, phkResult, NULL);

	if (lRet == ERROR_SUCCESS)
	{
		lRet = RegSetValueExW(
			(*phkResult),
			pszValueName, 0, REG_SZ,
			(LPBYTE)pszData,
			((DWORD)wcslen(pszData) + 1) * sizeof(WCHAR)
			);

		if (lRet != ERROR_SUCCESS)
		{
			RegCloseKey(*phkResult);
		}
	}

	return HRESULT_FROM_WIN32(lRet);
}

// Creates the registry entries for a COM object.

HRESULT RegisterObject(HMODULE hModule, const GUID& guid, const TCHAR *pszDescription, const TCHAR *pszThreadingModel)
{
	HKEY hKey = NULL;
	HKEY hSubkey = NULL;
	TCHAR achTemp[MAX_PATH];

	// Create the name of the key from the object's CLSID
	HRESULT hr = CreateObjectKeyName(guid, achTemp, MAX_PATH);

	// Create the new key.
	if (SUCCEEDED(hr))
	{
		hr = CreateRegKeyAndValue(HKEY_LOCAL_MACHINE, achTemp, NULL, pszDescription, &hKey);
	}

	if (SUCCEEDED(hr))
	{
		(void)GetModuleFileName(hModule, achTemp, MAX_PATH);

		hr = HRESULT_FROM_WIN32(GetLastError());
	}

	// Create the "InprocServer32" subkey
	if (SUCCEEDED(hr))
	{
		hr = CreateRegKeyAndValue(hKey, L"InProcServer32", NULL, achTemp, &hSubkey);
		RegCloseKey(hSubkey);
	}

	// Add a new value to the subkey, for "ThreadingModel" = <threading model>
	if (SUCCEEDED(hr))
	{
		hr = CreateRegKeyAndValue(hKey, L"InProcServer32", L"ThreadingModel", pszThreadingModel, &hSubkey);
		RegCloseKey(hSubkey);
	}

	// close hkeys
	RegCloseKey(hKey);
	return hr;
}

// Deletes the registry entries for a COM object.

HRESULT UnregisterObject(const GUID& guid)
{
	WCHAR achTemp[MAX_PATH];

	HRESULT hr = CreateObjectKeyName(guid, achTemp, MAX_PATH);
	if (SUCCEEDED(hr))
	{
		// Delete the key recursively.
		LONG lRes = RegDeleteTree(HKEY_LOCAL_MACHINE, achTemp);
		hr = HRESULT_FROM_WIN32(lRes);
	}
	return hr;
}
