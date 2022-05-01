
#if !defined(DebugInt3)
#	if !defined(_DEBUG) || defined(W64)
#		define DebugInt3 {}
#	else
#		define DebugInt3 {__asm {int 3}}
#	endif
#endif

#if !defined(__ATLCOMCLI_H__)
#	include <atlcomcli.h>
#endif

#define GENERIC_DEFINITIONS(class_name)									\
public:																	\
class_name() {}															\
class_name(const class_name& right)										\
{																		\
	obj = right.obj;													\
	creation_error.last_error = right.creation_error.last_error;		\
}																		\
void operator=(const class_name & right)								\
{																		\
	obj.Release();														\
	obj = right.obj;													\
	creation_error.last_error = right.creation_error.last_error;		\
}																		\
protected:																\
class_name(csp::Result &creation_error): BASE_CLASS(creation_error) {}	\
class_name(HRESULT creation_error): BASE_CLASS(creation_error) {}

#define CSP_PROP_1(name, type)											\
csp::Result put_##name(type x);											\
csp::Result get_##name(type &x);										\
__inline type _get_##name() { type x; get_##name(x); return x; }		\
__declspec(property(get=_get_##name, put=put_##name)) type name;

#define CSP_PROP_2(name, type)											\
csp::Result Set##name(type x);											\
csp::Result Get##name(type &x);											\
__inline type _Get##name() {type x; Get##name(x); return x; }			\
__declspec(property(get=_Get##name, put=Set##name)) type name;

#define CSP_QUERY_INTERFACE(sub_class_name)	sub_class_name Get##sub_class_name();

#define CSP_GET_OBJECT(owner_class, owner_method, return_class)			\
return_class owner_class::owner_method()								\
{																		\
	if(!IsValid()) return return_class(ERROR_INVALID_HANDLE);			\
	return_class p;														\
	p.creation_error = obj->owner_method(&p.obj);						\
	return p;															\
}

#define CSP_FUNC_1_PARAMETER(class_name, method, param_type)			\
csp::Result class_name::method(param_type& t)							\
{																		\
	CheckInitialized(class_name, method);								\
	return obj->method(&t);												\
}

namespace csp
{
	extern bool com_initialized;

	class TCHAR_TO_BSTR
	{
	public:
		BSTR conv;
		TCHAR_TO_BSTR(const TCHAR *input)
		{
#if defined(UNICODE)
			conv = SysAllocString(input);
#else
			conv = nullptr;
			int s = MultiByteToWideChar(CP_ACP, 0, input, -1, nullptr, 0);
			if(s == 0) return;
			s++;
			wchar_t *aux = new wchar_t[s];
			aux[0] = 0;
			try {
				if(MultiByteToWideChar(CP_ACP, 0, input, -1, aux, s) > 0) conv = SysAllocString(aux);
			}
			catch(...) {
			}
			delete aux;
#endif
		}
		~TCHAR_TO_BSTR() { SysFreeString(conv); }
		operator BSTR() const { return conv; }
	};

	class Result
	{
	public:
		HRESULT last_error;
		__inline Result()
		{
			last_error = S_OK;
		}

		__inline Result(HRESULT le)
		{
			if (le > 0)
			{
				//DebugInt3;
			}
			last_error = le;
		}

		__inline Result(const Result &r)
		{
			last_error = r.last_error;
		}

		__inline bool Succeeded() const
		{
			return last_error == S_OK;
		}
		
		__inline bool operator !() const
		{
			return last_error != S_OK;
		}
	};

	template<class com_class>
	class Base
	{
	public:
		typedef com_class CSP_INTERFACE;
		typedef csp::Base<com_class> BASE_CLASS;

		Base()
		{
			creation_error.last_error = S_OK;
			if(com_initialized) return;
			com_initialized = true;
			CoInitialize(NULL);
		}

		template <class cast> cast Cast()
		{
			cast dest;
			if (!obj) return dest;

			cast::CSP_INTERFACE* c;
			if (!SUCCEEDED(obj->QueryInterface(__uuidof(cast::CSP_INTERFACE), (void**)&c)) || !c) return dest;

			class x : public cast
			{
			public:
				void Set(cast::CSP_INTERFACE* y)
				{
					obj.Attach(y);
				}
			} u;
			u.Set(c);
			return u;
		}

		bool IsValid() const
		{
			return obj != nullptr;
		}

		void Release()
		{
			obj.Release();
		}

		operator bool() const
		{
			return obj != nullptr;
		}

		bool operator !() const
		{
			return !obj;
		}

		Result& CreationError() const { return IsValid()?ERROR_SUCCESS:creation_error; }

		CComPtr<com_class> obj;
	protected:

		Base(Result &creation_error)
		{
			this->creation_error.last_error = creation_error.last_error;
		}

		Base(HRESULT creation_error)
		{
			this->creation_error.last_error = creation_error;
		}

		Result creation_error;
	};

	template<class csp_item_class, typename com_enumerator_interface>
	class iterator
	{
	public:
		__inline iterator()
		{
		}

		__inline iterator(const iterator& i)
		{
			if (!i.pEm) return;
			pEm = i.pEm;
			ref.obj = i.ref.obj;
		}

		__inline void operator=(const iterator& i)
		{
			if (!i.pEm) return;
			pEm = i.pEm;
			ref.obj = i.ref.obj;
		}

		__inline iterator(const CComPtr<typename com_enumerator_interface> &cpy)
		{
			if (!cpy) return;// cpy could be nullptr, for example: calls to CreateClassEnumerator with an valid category, but with no devices in that category
			pEm = cpy;
			ULONG cFetched;
			HRESULT hr = pEm->Next(1, &ref.obj, &cFetched);
			if (!SUCCEEDED(hr) || cFetched != 1) ref.obj.Release();
		}

		__inline csp_item_class& operator*()
		{
			return ref;
		}

		__inline bool operator !=(const iterator &right) const
		{
			return ref.obj != right.ref.obj;
		}

		__inline void operator++()
		{
			ref.obj.Release();
			if (!pEm) return;
			ULONG cFetched;
			if (!SUCCEEDED(pEm->Next(1, &ref.obj, &cFetched)) || cFetched != 1) ref.obj.Release();
		}

	private:
		static bool com_initialized;
		csp_item_class ref;
		CComPtr<typename com_enumerator_interface> pEm;
	};

	template<class csp_item, typename com_enumerator_interface>
	class Collection: public Base<com_enumerator_interface>
	{
	public:
		typedef csp::Collection<csp_item, com_enumerator_interface> BASE_CLASS;
		friend class csp::iterator<csp_item, com_enumerator_interface>;

		csp::iterator<csp_item, com_enumerator_interface> begin() const
		{
			if (!IsValid()) return csp::iterator<csp_item, com_enumerator_interface>();
			return csp::iterator<csp_item, com_enumerator_interface>(obj);
		}

		csp::iterator<csp_item, com_enumerator_interface>& end() const
		{
			static csp::iterator<csp_item, com_enumerator_interface> i;
			return i;
		}

		Collection() {}

	protected:
		Collection(HRESULT r): Base(r) {}
		Collection(csp::Result r): Base(r) {}
	};

	class LastError
	{
	public:
		HRESULT last_error;
		LastError() { last_error = S_OK; }
	};

	class Unknown : public Base<IUnknown>
	{
	public:
		Unknown(CComPtr<IUnknown> &external_obj) { obj = external_obj; }
		const CComPtr<IUnknown>& getObj() const { return obj; }
		GENERIC_DEFINITIONS(Unknown)
	};
}

#if !defined(Trace2)
#define Trace2(context_class, context_method, message)				/*{ const type_info &__i = typeid(*this); LPCSTR name = __i.name(); TraceObject.WriteMsg(__FILE__, __LINE__, name?name:#context_class, #context_method, #message, (LPCTSTR)NULL); }*/
#endif
#define CheckArguments(cond, context_class, context_method)	{if(cond) {Trace2(context_class, context_method, "inval. argments"); DebugInt3; return E_INVALIDARG;}}
#define CheckInitialized(context_class, context_method)		{if(!IsValid()) {Trace2(context_class, context_method, "!init"); DebugInt3; return E_HANDLE;}}

#define CSP_PROP_1_IMP(class_name, prop_name, type)			\
csp::Result class_name::put_##prop_name(type x)				\
{															\
	CheckInitialized(class_name, put_##prop_name);			\
	try {													\
		return obj->put_##prop_name(x);						\
	}														\
	catch(...) {											\
		DebugInt3;											\
		return E_POINTER;									\
	}														\
}															\
csp::Result class_name::get_##prop_name(type &x)			\
{															\
	CheckInitialized(class_name, get_##prop_name);			\
	try {													\
		return obj->get_##prop_name(&x);					\
	}														\
	catch(...) {											\
		DebugInt3;											\
		return E_POINTER;									\
	}														\
}

#define CSP_PROP_2_IMP(class_name, prop_name, type)			\
csp::Result class_name::Set##prop_name(type x)				\
{															\
	CheckInitialized(class_name, Set##prop_name);			\
	try {													\
		return obj->Set##prop_name(x);						\
	}														\
	catch(...) {											\
		DebugInt3;											\
		return E_POINTER;									\
	}														\
}															\
csp::Result class_name::Get##prop_name(type &x)				\
{															\
	CheckInitialized(class_name, Get##prop_name);			\
	try {													\
		return obj->Get##prop_name(&x);						\
	}														\
	catch(...) {											\
		DebugInt3;											\
		return E_POINTER;									\
	}														\
}

#define CSP_QUERY_INTERFACE_IMP(class_name, sub_class_name)	\
sub_class_name class_name::Get##sub_class_name()			\
{															\
	sub_class_name mc;										\
	if (!IsValid()) return mc;								\
	try {													\
	if (SUCCEEDED(obj->QueryInterface(__uuidof(sub_class_name::CSP_INTERFACE), (void **)&mc.obj))) return mc;\
	}														\
	catch(...) {											\
		DebugInt3;											\
	}														\
	return sub_class_name(E_POINTER);						\
}

