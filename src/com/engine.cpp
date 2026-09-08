#include <blipbridge/dispatch.hpp>
#include <atomic>
#include <map>
#include <filesystem>
using namespace bb;
extern int runExperiment(int,wchar_t**);
extern HRESULT traceUserPicture(IDispatch*,const std::wstring&);
extern HRESULT focusedBenchmarks(IDispatch*,const std::wstring&);
extern HRESULT memoryFillExperiment(IDispatch*,SAFEARRAY*,const std::wstring&);
extern HRESULT stressExperiment(IDispatch*,const std::wstring&);
extern HRESULT traceCachedApply(IDispatch*,IDispatch*,const std::wstring&);
static HMODULE module;
static std::atomic<long> objects{0},locks{0};
// BlipBridge.Engine: {2E2E2731-C523-486B-89CB-2A89484F1E32}
static const CLSID clsid={0x2e2e2731,0xc523,0x486b,{0x89,0xcb,0x2a,0x89,0x48,0x4f,0x1e,0x32}};
static const IID extiid={0xb65ad801,0xabaf,0x11d0,{0xbb,0x8b,0x00,0xa0,0xc9,0x0f,0x27,0x44}};
struct Ext:IDispatch {
 virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*,long,IDispatch*,SAFEARRAY**)=0;
 virtual HRESULT STDMETHODCALLTYPE OnDisconnection(long,SAFEARRAY**)=0;
 virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**)=0;
 virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**)=0;
 virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**)=0;
};
class Engine final:public Ext {
 std::atomic<ULONG> refs{1};DWORD thread=GetCurrentThreadId();std::map<long,Value> textures;long next=1;std::wstring lastError;
public:
 Engine(){++objects;OutputDebugStringW(L"BlipBridge 0.1: PickupApplyFallback\n");}
 ~Engine(){--objects;}
 HRESULT STDMETHODCALLTYPE QueryInterface(REFIID i,void** p)override{if(!p)return E_POINTER;*p=nullptr;if(i==IID_IUnknown||i==IID_IDispatch||i==extiid){*p=static_cast<Ext*>(this);AddRef();return S_OK;}return E_NOINTERFACE;}
 ULONG STDMETHODCALLTYPE AddRef()override{return ++refs;}
 ULONG STDMETHODCALLTYPE Release()override{auto n=--refs;if(!n)delete this;return n;}
 HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* n)override{if(!n)return E_POINTER;*n=0;return S_OK;}
 HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT,LCID,ITypeInfo**)override{return E_NOTIMPL;}
 HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID iid,LPOLESTR* ns,UINT n,LCID,DISPID* ids)override{
  if(iid!=IID_NULL)return DISP_E_UNKNOWNINTERFACE;
  static const wchar_t* names[]={L"GetVersion",L"GetBackendName",L"GetCapabilities",L"RegisterTextureShape",L"ApplyTexture",L"ReleaseTexture",L"ClearTextures",L"GetTextureCount",L"GetLastError",L"LoadTexture",L"SetImageBytes",L"RunBenchmarks",L"GetHostProcessId",L"TraceUserPicture",L"RunFocusedBenchmarks",L"MemoryFillExperiment",L"RunStress",L"TraceCachedApply"};
  for(UINT j=0;j<n;j++){ids[j]=DISPID_UNKNOWN;for(int k=0;k<18;k++)if(!_wcsicmp(ns[j],names[k]))ids[j]=k+1;if(ids[j]==DISPID_UNKNOWN)return DISP_E_UNKNOWNNAME;}return S_OK;
 }
 HRESULT STDMETHODCALLTYPE Invoke(DISPID id,REFIID iid,LCID,WORD flags,DISPPARAMS* p,VARIANT* out,EXCEPINFO* ex,UINT*)override{
  if(iid!=IID_NULL)return DISP_E_UNKNOWNINTERFACE;
  if(!(flags&(DISPATCH_METHOD|DISPATCH_PROPERTYGET)))return DISP_E_MEMBERNOTFOUND;
  if(GetCurrentThreadId()!=thread)return RPC_E_WRONG_THREAD;
  try{
   Value r;auto arg=[&](UINT i){if(!p||i>=p->cArgs)throw Error(DISP_E_BADPARAMCOUNT,"Missing argument");Value v;check(VariantCopyInd(&v.v,&p->rgvarg[p->cArgs-1-i]),"Argument");return v;};
   auto count=[&](UINT n){if(!p||p->cArgs!=n)throw Error(DISP_E_BADPARAMCOUNT,"Wrong argument count");};
   switch(id){
    case 1:count(0);r=Value(L"0.1.0-experimental");break;
    case 2:count(0);r=Value(L"PickupApplyFallback");break;
    case 3:count(0);r=Value(L"MemoryImageToFill=False;CachedTextureApply=False;PickUpFallback=True;FillOnly=False;InternalBackend=False");break;
    case 4:{count(1);auto s=arg(0);auto t=get(s.obj(),L"Type").integer();if(t!=1&&t!=5)throw Error(E_INVALIDARG,"Expected AutoShape or Freeform donor");if(get(get(s.obj(),L"Fill").obj(),L"Type").integer()!=6)throw Error(E_INVALIDARG,"Donor must have picture fill");if(next==LONG_MAX)throw Error(E_OUTOFMEMORY,"Handle space exhausted");long h=next++;textures.emplace(h,std::move(s));r=Value(h);break;}
    case 5:{count(2);auto s=arg(0);long h=arg(1).integer();auto it=textures.find(h);if(it==textures.end())throw Error(HRESULT(0x80040206),"Texture handle not found");auto type=get(s.obj(),L"Type").integer();if(type!=1&&type!=5)throw Error(E_INVALIDARG,"Target must be AutoShape or Freeform");call(it->second.obj(),L"PickUp");call(s.obj(),L"Apply");r=Value(-1L);break;}
    case 6:count(1);if(!textures.erase(arg(0).integer()))throw Error(HRESULT(0x80040206),"Texture handle not found");break;
    case 7:count(0);textures.clear();break;
    case 8:count(0);r=Value((long)textures.size());break;
    case 9:count(0);r=Value(lastError.c_str());break;
    case 10:case 11:throw Error(E_NOTIMPL,"No validated memory-image backend on this Office build. RegisterTextureShape is the explicit fallback.");
    case 12:{count(1);if(!GetModuleHandleW(L"POWERPNT.EXE"))throw Error(E_ACCESSDENIED,"RunBenchmarks must run inside PowerPoint via research add-in");auto root=arg(0).str();wchar_t name[]=L"bb";wchar_t* a[]={name,root.data()};r=Value((long)runExperiment(2,a));break;}
    case 13:count(0);r=Value((long)GetCurrentProcessId());break;
    case 14:count(2);if(!GetModuleHandleW(L"POWERPNT.EXE"))throw Error(E_ACCESSDENIED,"Trace requires PowerPoint host");check(traceUserPicture(arg(0).obj(),arg(1).str()),"TraceUserPicture");break;
    case 15:count(2);if(!GetModuleHandleW(L"POWERPNT.EXE"))throw Error(E_ACCESSDENIED,"Requires PowerPoint host");check(focusedBenchmarks(arg(0).obj(),arg(1).str()),"Focused benchmarks");break;
    case 16:{count(3);if(!GetModuleHandleW(L"POWERPNT.EXE"))throw Error(E_ACCESSDENIED,"Requires PowerPoint host");auto bytes=arg(1);if(bytes.v.vt!=(VT_ARRAY|VT_UI1))throw Error(E_INVALIDARG,"Expected Byte array");check(memoryFillExperiment(arg(0).obj(),bytes.v.parray,arg(2).str()),"Memory fill experiment");break;}
    case 17:count(2);if(!GetModuleHandleW(L"POWERPNT.EXE"))throw Error(E_ACCESSDENIED,"Requires PowerPoint host");check(stressExperiment(arg(0).obj(),arg(1).str()),"Stress experiment");break;
    case 18:count(3);if(!GetModuleHandleW(L"POWERPNT.EXE"))throw Error(E_ACCESSDENIED,"Requires PowerPoint host");check(traceCachedApply(arg(0).obj(),arg(1).obj(),arg(2).str()),"Cached trace");break;
    default:return DISP_E_MEMBERNOTFOUND;
   }
   if(id!=9)lastError.clear();
   if(out){VariantInit(out);check(VariantCopy(out,&r.v),"Return");}return S_OK;
  }catch(const Error& e){lastError.assign(e.what(),e.what()+strlen(e.what()));if(ex){*ex={};ex->bstrSource=SysAllocString(L"BlipBridge.Engine");ex->bstrDescription=SysAllocString(lastError.c_str());ex->scode=e.hr;return DISP_E_EXCEPTION;}return e.hr;}catch(const std::bad_alloc&){return E_OUTOFMEMORY;}catch(...){return E_UNEXPECTED;}
 }
 HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*,long,IDispatch* addin,SAFEARRAY**)override{
  try{put(addin,L"Object",Value(static_cast<IDispatch*>(this)));return S_OK;}catch(const Error& e){return e.hr;}catch(...){return E_UNEXPECTED;}
 }
 HRESULT STDMETHODCALLTYPE OnDisconnection(long,SAFEARRAY**)override{textures.clear();return S_OK;}
 HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**)override{return S_OK;}
 HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**)override{return S_OK;}
 HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**)override{textures.clear();return S_OK;}
};
class Factory final:public IClassFactory {
 std::atomic<ULONG> refs{1};
public:
 Factory(){++objects;}~Factory(){--objects;}
 HRESULT STDMETHODCALLTYPE QueryInterface(REFIID i,void** p)override{if(!p)return E_POINTER;*p=nullptr;if(i==IID_IUnknown||i==IID_IClassFactory){*p=this;AddRef();return S_OK;}return E_NOINTERFACE;}
 ULONG STDMETHODCALLTYPE AddRef()override{return ++refs;}ULONG STDMETHODCALLTYPE Release()override{auto n=--refs;if(!n)delete this;return n;}
 HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer,REFIID i,void** p)override{if(!p)return E_POINTER;*p=nullptr;if(outer)return CLASS_E_NOAGGREGATION;try{auto e=new Engine;auto h=e->QueryInterface(i,p);e->Release();return h;}catch(...){return E_OUTOFMEMORY;}}
 HRESULT STDMETHODCALLTYPE LockServer(BOOL b)override{b?++locks:--locks;return S_OK;}
};
extern "C" __declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID c,REFIID i,void** p){if(!p)return E_POINTER;*p=nullptr;if(c!=clsid)return CLASS_E_CLASSNOTAVAILABLE;try{auto f=new Factory;auto h=f->QueryInterface(i,p);f->Release();return h;}catch(...){return E_OUTOFMEMORY;}}
extern "C" __declspec(dllexport) HRESULT __stdcall DllCanUnloadNow(){return objects==0&&locks==0?S_OK:S_FALSE;}
BOOL WINAPI DllMain(HINSTANCE h,DWORD reason,LPVOID){if(reason==DLL_PROCESS_ATTACH){module=h;DisableThreadLibraryCalls(h);}return TRUE;}
