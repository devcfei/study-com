#include "stdafx.h"
#include "comaimpl.h"



HRESULT CComa::CreateInstance(REFIID iid, void **ppComa)
{
	*ppComa = NULL;

	HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

	if (IsEqualGUID(iid, __uuidof(IComa)))
	{
		auto sObject = new CComa;
		if (sObject != nullptr)
		{
			*ppComa = sObject;
			hr = S_OK;

		}
	}

	return hr;
}


//IUnknown methods
IFACEMETHODIMP  CComa::QueryInterface(REFIID riid, void **ppvObject)
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
	else if (riid == __uuidof(IComa) /*IID_IComa*/)
	{
		*ppvObject = (IComa*)this;
	}
	else
	{
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}

	AddRef();

	return hr;
}

IFACEMETHODIMP_(ULONG)  CComa::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) CComa::Release()
{
	long cRef = InterlockedDecrement(&m_cRef);
	if (cRef == 0)
	{
		delete this;
	}
	return cRef;
}




STDMETHODIMP CComa::SetVal(UINT32 val)
{

	m_val = val;
	return S_OK;
}

STDMETHODIMP CComa::GetVal(UINT32 *pVal)
{
	HRESULT hr = S_OK;


	if (!pVal)
	{
		return E_POINTER;
	}
	*pVal = m_val;

	return hr;
}