#include <blipbridge/dispatch.hpp>
#include <filesystem>
#include <fstream>
#include <numeric>
#include <psapi.h>
using namespace bb;
// Bind through the actual dual-interface typelib, never an Office internal address.
struct Bound {
 IUnknown* object=nullptr;void* function=nullptr;
 Bound(IDispatch* d,const wchar_t* name){
  ITypeInfo* ti=nullptr;check(d->GetTypeInfo(0,0,&ti),"type info");TYPEATTR* ta=nullptr;ti->GetTypeAttr(&ta);GUID iid=ta->guid;ti->ReleaseTypeAttr(ta);
  check(d->QueryInterface(iid,(void**)&object),"dual interface");DISPID id;auto n=const_cast<wchar_t*>(name);check(ti->GetIDsOfNames(&n,1,&id),"member");
  ti->GetTypeAttr(&ta);for(UINT i=0;i<ta->cFuncs;i++){FUNCDESC* f=nullptr;ti->GetFuncDesc(i,&f);if(f->memid==id&&f->invkind==INVOKE_FUNC)function=(*(void***)object)[f->oVft/sizeof(void*)];ti->ReleaseFuncDesc(f);}
  ti->ReleaseTypeAttr(ta);ti->Release();if(!function)throw Error(E_NOINTERFACE,"No dual method");
 }
 ~Bound(){if(object)object->Release();}
 void noargs(){using F=HRESULT(STDMETHODCALLTYPE*)(IUnknown*);check(((F)function)(object),"dual call");}
 void picture(BSTR s){using F=HRESULT(STDMETHODCALLTYPE*)(IUnknown*,BSTR);check(((F)function)(object,s),"dual UserPicture");}
};
static long long tick(){LARGE_INTEGER n;QueryPerformanceCounter(&n);return n.QuadPart;}
static size_t memory(){PROCESS_MEMORY_COUNTERS_EX p{};GetProcessMemoryInfo(GetCurrentProcess(),(PROCESS_MEMORY_COUNTERS*)&p,sizeof p);return p.PrivateUsage;}
template<class F>void measure(std::ofstream& log,const char* label,int n,F f){
 LARGE_INTEGER q;QueryPerformanceFrequency(&q);std::vector<double> times;for(int i=0;i<10;i++)f(i);auto before=memory();
 for(int i=0;i<n;i++){auto t=tick();f(i);times.push_back((tick()-t)*1e6/q.QuadPart);}auto after=memory();
 auto total=std::accumulate(times.begin(),times.end(),0.);std::sort(times.begin(),times.end());log<<label<<','<<n<<','<<total/1000<<','<<total/n<<','<<times[n/2]<<','<<times.front()<<','<<times.back()<<','<<n*1e6/total<<','<<before<<','<<after<<'\n';log.flush();
}
HRESULT focusedBenchmarks(IDispatch* app,const std::wstring& root){
 try{
 auto presentations=get(app,L"Presentations");auto pres=call(presentations.obj(),L"Add",{Value(0L)});
 try{
 auto slide=call(get(pres.obj(),L"Slides").obj(),L"Add",{Value(1L),Value(12L)});auto shapes=get(slide.obj(),L"Shapes");
 auto rect=[&](){return call(shapes.obj(),L"AddShape",{Value(1L),Value(10.),Value(10.),Value(80.),Value(60.)});};
 auto target=rect(),donor=rect();auto tf=get(target.obj(),L"Fill"),df=get(donor.obj(),L"Fill");
 Value p0((std::filesystem::path(root)/L"artifacts/textures/texture_64_0.png").c_str());Value p1((std::filesystem::path(root)/L"artifacts/textures/texture_64_1.png").c_str());call(df.obj(),L"UserPicture",{p0});
 Bound user(tf.obj(),L"UserPicture"),pickup(donor.obj(),L"PickUp"),apply(target.obj(),L"Apply");
 std::ofstream log(std::filesystem::path(root)/L"artifacts/focused.csv");log<<"operation,count,total_ms,average_us,median_us,min_us,max_us,ops_per_sec,private_before,private_after\n";
 measure(log,"Hidden_Invoke_UserPicture",1000,[&](int){call(tf.obj(),L"UserPicture",{p0});});
 measure(log,"Hidden_Dual_UserPicture_same",1000,[&](int){user.picture(p0.v.bstrVal);});
 measure(log,"Hidden_Dual_UserPicture_alternate",1000,[&](int i){user.picture(i%2?p0.v.bstrVal:p1.v.bstrVal);});
 pickup.noargs();measure(log,"Hidden_Dual_Apply",10000,[&](int){apply.noargs();});
 measure(log,"Hidden_Dual_PickUp_Apply",10000,[&](int){pickup.noargs();apply.noargs();});
 measure(log,"Hidden_Invoke_PickUp_Apply",1000,[&](int){call(donor.obj(),L"PickUp");call(target.obj(),L"Apply");});
 call(pres.obj(),L"SaveAs",{Value((std::filesystem::path(root)/L"artifacts/focused.pptx").c_str()),Value(24L)});
 }catch(...){put(pres.obj(),L"Saved",Value(-1L));call(pres.obj(),L"Close");throw;}
 call(pres.obj(),L"Close");return S_OK;
 }catch(const Error& e){OutputDebugStringA(e.what());return e.hr;}catch(...){return E_UNEXPECTED;}
}
