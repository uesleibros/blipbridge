// Isolated, reversible IAT instrumentation. Never enabled by normal Engine calls.
#include <windows.h>
#include <psapi.h>
#include <fstream>
#include <filesystem>
#include <vector>
#include <string>
#include <mutex>
#include <blipbridge/dispatch.hpp>
namespace {
std::ofstream logFile;
std::mutex logMutex;
HANDLE tracked=INVALID_HANDLE_VALUE;
DWORD traceThread=0;
using OpenFn=decltype(&CreateFileW);using ReadFn=decltype(&ReadFile);using CloseFn=decltype(&CloseHandle);
OpenFn realOpen=CreateFileW;ReadFn realRead=ReadFile;CloseFn realClose=CloseHandle;
void stack(const char* event){
 std::lock_guard lock(logMutex);logFile<<"\nEVENT "<<event<<" thread="<<GetCurrentThreadId()<<'\n';
 void* frames[48];USHORT n=CaptureStackBackTrace(1,48,frames,nullptr);
 for(USHORT i=0;i<n;i++){HMODULE m=nullptr;wchar_t path[32768]{};GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS|GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,(LPCWSTR)frames[i],&m);GetModuleFileNameW(m,path,32768);auto name=std::filesystem::path(path).filename().string();logFile<<name<<"+0x"<<std::hex<<((uintptr_t)frames[i]-(uintptr_t)m)<<" absolute=0x"<<(uintptr_t)frames[i]<<std::dec<<'\n';}logFile.flush();
}
HANDLE WINAPI openHook(LPCWSTR name,DWORD access,DWORD share,LPSECURITY_ATTRIBUTES sa,DWORD disposition,DWORD flags,HANDLE templ){
 bool match=GetCurrentThreadId()==traceThread&&name&&wcsstr(name,L"BB_TRACE_UNIQUE_TEXTURE_01");
 if(match)stack("CreateFileW entry (unique texture)");
 HANDLE h=realOpen(name,access,share,sa,disposition,flags,templ);DWORD error=GetLastError();
 if(match){tracked=h;logFile<<"CreateFileW result="<<h<<" access="<<access<<" flags="<<flags<<" error="<<error<<'\n';logFile.flush();}SetLastError(error);return h;
}
BOOL WINAPI readHook(HANDLE h,LPVOID p,DWORD n,LPDWORD got,LPOVERLAPPED ov){
 if(h==tracked&&GetCurrentThreadId()==traceThread){stack("ReadFile unique texture");logFile<<"requested="<<n<<'\n';}
 return realRead(h,p,n,got,ov);
}
BOOL WINAPI closeHook(HANDLE h){if(h==tracked&&GetCurrentThreadId()==traceThread){stack("CloseHandle unique texture");tracked=INVALID_HANDLE_VALUE;}return realClose(h);}
struct Patch {uintptr_t* slot;uintptr_t original;uintptr_t replacement;};
std::vector<Patch> patches;
void change(uintptr_t* slot,uintptr_t expected,uintptr_t replacement){
 if(*slot!=expected)throw bb::Error(E_FAIL,"IAT invariant failed");
 DWORD old=0;if(!VirtualProtect(slot,sizeof(*slot),PAGE_READWRITE,&old))throw bb::Error(HRESULT_FROM_WIN32(GetLastError()),"IAT protection");
 InterlockedExchangePointer((void* volatile*)slot,(void*)replacement);DWORD ignored;VirtualProtect(slot,sizeof(*slot),old,&ignored);
}
void restore(){for(auto i=patches.rbegin();i!=patches.rend();++i){if(*i->slot==i->replacement)change(i->slot,i->replacement,i->original);}patches.clear();}
void install(){
 HMODULE mods[1024];DWORD bytes;bb::check(EnumProcessModules(GetCurrentProcess(),mods,sizeof(mods),&bytes)?S_OK:E_FAIL,"Modules");
 for(unsigned i=0;i<bytes/sizeof(HMODULE)&&i<1024;i++){
  wchar_t path[32768]{};GetModuleFileNameW(mods[i],path,32768);std::wstring lower=path;for(auto& c:lower)c=towlower(c);
  if(lower.find(L"microsoft office")==std::wstring::npos && lower.find(L"microsoft shared\\office")==std::wstring::npos)continue;
  auto base=(BYTE*)mods[i];auto dos=(IMAGE_DOS_HEADER*)base;if(dos->e_magic!=IMAGE_DOS_SIGNATURE)continue;
  auto nt=(IMAGE_NT_HEADERS64*)(base+dos->e_lfanew);if(nt->Signature!=IMAGE_NT_SIGNATURE||nt->FileHeader.Machine!=IMAGE_FILE_MACHINE_AMD64)continue;
  auto size=nt->OptionalHeader.SizeOfImage;auto dir=nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];if(!dir.VirtualAddress||dir.VirtualAddress+dir.Size>size)continue;
  auto desc=(IMAGE_IMPORT_DESCRIPTOR*)(base+dir.VirtualAddress);
  for(; (BYTE*)(desc+1)<=base+dir.VirtualAddress+dir.Size&&desc->Name;desc++){
   if(!desc->FirstThunk||desc->FirstThunk>=size)continue;
   auto slot=(uintptr_t*)(base+desc->FirstThunk);
   for(;(BYTE*)(slot+1)<=base+size&&*slot;slot++){
    uintptr_t replacement=0;if(*slot==(uintptr_t)realOpen)replacement=(uintptr_t)openHook;
    if(*slot==(uintptr_t)realRead)replacement=(uintptr_t)readHook;if(*slot==(uintptr_t)realClose)replacement=(uintptr_t)closeHook;
    if(replacement){logFile<<"IAT "<<std::filesystem::path(path).filename().string()<<" slotRVA=0x"<<std::hex<<((BYTE*)slot-base)<<std::dec<<'\n';patches.push_back({slot,*slot,replacement});change(slot,*slot,replacement);}
   }
  }
 }
 logFile<<"Patched slots="<<patches.size()<<'\n';logFile.flush();
}
}
HRESULT traceUserPicture(IDispatch* fill,const std::wstring& root){
 try{
 logFile.open(std::filesystem::path(root)/L"artifacts/userpicture_trace.txt");traceThread=GetCurrentThreadId();install();
 try {bb::call(fill,L"UserPicture",{bb::Value((std::filesystem::path(root)/L"artifacts/textures/BB_TRACE_UNIQUE_TEXTURE_01.png").c_str())});}catch(...){restore();throw;}
 restore();logFile<<"All IAT patches restored\n";logFile.close();return S_OK;
 }catch(const bb::Error& e){restore();logFile<<e.what()<<'\n';logFile.close();return e.hr;}catch(...){restore();logFile.close();return E_UNEXPECTED;}
}
