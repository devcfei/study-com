#include <Windows.h>
#include <Shlwapi.h>
#include <iostream>

#include <objbase.h>
#include <initguid.h>
#include "coma.h"
#include "comaclsid.h"


HRESULT ComRegister(LPCTSTR lpszPath)
{
	HRESULT hr =E_FAIL;
	typedef HRESULT(WINAPI * FREG)();

	HMODULE hDLL = ::LoadLibrary(lpszPath);
	if (hDLL)
	{
		FREG lpfunc = (FREG)::GetProcAddress(hDLL, TEXT("DllRegisterServer"));

		if (lpfunc)	
			hr = lpfunc();
		::FreeLibrary(hDLL);
	}

	//::SetCurrentDirectory(szWorkPath);

	return hr;
}

HRESULT ComUnregister(LPCTSTR lpszPath)
{
	HRESULT hr = E_FAIL;
	typedef HRESULT(WINAPI * FREG)();

	HMODULE hDLL = ::LoadLibrary(lpszPath);
	if (hDLL)
	{
		FREG lpfunc = (FREG)::GetProcAddress(hDLL, TEXT("DllUnregisterServer"));

		if (lpfunc)
			hr = lpfunc();

		::FreeLibrary(hDLL);
	}
	return hr;

}


int _cdecl main(int argc, LPCTSTR argv[])
{
	HRESULT hr;

	if (argc < 2)
	{
		std::cout << "need a argument for dll path!" << std::endl;
		return -1;
	}


	hr = CoInitializeEx(NULL,0);
	if (FAILED(hr))
	{
		return hr;
	}

	hr = ComRegister(argv[1]);
	if (FAILED(hr))
	{
		std::cout << "COM register failed, are you Administrator?" << std::endl;
		return -1;
	}
	IComa *piComa;

	hr = ::CoCreateInstance(CLSID_Coma,NULL,
		CLSCTX_INPROC_SERVER,
		__uuidof(IComa),//IID_IComa,
		(void **)&piComa);

	if (SUCCEEDED(hr))
	{
		UINT32 val = 0;
		hr = piComa->SetVal(10);
		
		val = 0;
		hr = piComa->GetVal(&val);

		if (val != 10)
		{
			std::cout << "val incorrect!" << std::endl;
		}

		std::cout << "val =" << val << std::endl;


		piComa->Release();

	}


	hr = ComUnregister(argv[1]);
	if (FAILED(hr))
	{
		std::cout << "COM register failed, are you Administrator?" << std::endl;
		return -1;
	}
	CoUninitialize();

	return 0;
}
