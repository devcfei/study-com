#pragma once

#include "coma.h"




class CComa : public IComa
{
public:
	static HRESULT CreateInstance(REFIID iid, void **ppComa);

	CComa():m_cRef(1),
		m_val(0)
	{

	}


	STDMETHOD(SetVal)(UINT32 val);
	STDMETHOD(GetVal)(UINT32 *pVal);


	//IUnknown methods
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	IFACEMETHODIMP_(ULONG) AddRef();	
	IFACEMETHODIMP_(ULONG) Release();
	


protected:
	UINT32 m_val;

	long m_cRef;

};
