#pragma once
#include <windows.h>
#include <oleauto.h>
#include <stdexcept>
#include <string>
#include <vector>
#include <algorithm>
namespace bb {
struct Error:std::runtime_error { HRESULT hr; Error(HRESULT h,const std::string& s):runtime_error(s),hr(h){} };
inline void check(HRESULT hr,const char* s){if(FAILED(hr))throw Error(hr,s);}
struct Value {
 VARIANT v;
 Value(){VariantInit(&v);} explicit Value(long x):Value(){v.vt=VT_I4;v.lVal=x;}
 explicit Value(double x):Value(){v.vt=VT_R8;v.dblVal=x;}
 explicit Value(const wchar_t* x):Value(){v.vt=VT_BSTR;v.bstrVal=SysAllocString(x);}
 explicit Value(IDispatch* x):Value(){v.vt=VT_DISPATCH;v.pdispVal=x;if(x)x->AddRef();}
 Value(const Value& x):Value(){check(VariantCopy(&v,const_cast<VARIANT*>(&x.v)),"VariantCopy");}
 Value(Value&& x) noexcept:v(x.v){VariantInit(&x.v);}
 Value& operator=(Value x){std::swap(v,x.v);return *this;}
 ~Value(){VariantClear(&v);}
 IDispatch* obj()const {if(v.vt!=VT_DISPATCH || !v.pdispVal)throw Error(E_NOINTERFACE,"Expected dispatch");return v.pdispVal;}
 long integer()const{Value r;check(VariantChangeType(&r.v,const_cast<VARIANT*>(&v),0,VT_I4),"integer");return r.v.lVal;}
 double number()const{Value r;check(VariantChangeType(&r.v,const_cast<VARIANT*>(&v),0,VT_R8),"number");return r.v.dblVal;}
 std::wstring str()const{Value r;check(VariantChangeType(&r.v,const_cast<VARIANT*>(&v),0,VT_BSTR),"string");return r.v.bstrVal?std::wstring(r.v.bstrVal,SysStringLen(r.v.bstrVal)):L"";}
};
inline Value invoke(IDispatch* o,const wchar_t* name,WORD flags,std::vector<Value> a={}){
 DISPID id;auto n=const_cast<wchar_t*>(name);check(o->GetIDsOfNames(IID_NULL,&n,1,LOCALE_USER_DEFAULT,&id),"GetIDsOfNames");
 std::reverse(a.begin(),a.end());std::vector<VARIANT> raw;for(auto& x:a)raw.push_back(x.v);
 DISPID put=DISPID_PROPERTYPUT;DISPPARAMS p{raw.data(),nullptr,(UINT)raw.size(),0};if(flags&DISPATCH_PROPERTYPUT){p.rgdispidNamedArgs=&put;p.cNamedArgs=1;}
 Value r;EXCEPINFO ex{};UINT bad{};HRESULT hr=o->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,flags,&p,&r.v,&ex,&bad);
 std::string msg="Invoke ";while(*name)msg.push_back((char)*name++);
 if(ex.bstrDescription){msg+=" : ";for(auto c:std::wstring(ex.bstrDescription))msg.push_back(c<128?(char)c:'?');}
 SysFreeString(ex.bstrSource);SysFreeString(ex.bstrDescription);SysFreeString(ex.bstrHelpFile);check(hr,msg.c_str());return r;
}
inline Value get(IDispatch* o,const wchar_t* n){return invoke(o,n,DISPATCH_PROPERTYGET);}
inline Value call(IDispatch* o,const wchar_t* n,std::vector<Value> a={}){return invoke(o,n,DISPATCH_METHOD,std::move(a));}
inline void put(IDispatch* o,const wchar_t* n,Value x){invoke(o,n,DISPATCH_PROPERTYPUT,{std::move(x)});}
inline Value item(IDispatch* o,long i){return invoke(o,L"Item",DISPATCH_METHOD|DISPATCH_PROPERTYGET,{Value(i)});}
}

