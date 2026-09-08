#include <blipbridge/dispatch.hpp>
#include <oleidl.h>
#include <ocidl.h>
#include <servprov.h>
#include <filesystem>
#include <fstream>
#include <iostream>
using namespace bb;
static void probe(IDispatch* object,const char* label,std::ofstream& out){
 struct Entry{const char* name;IID id;};
 const Entry entries[]={{"IStream",IID_IStream},{"ILockBytes",IID_ILockBytes},{"IPersist",IID_IPersist},{"IPersistStream",IID_IPersistStream},{"IPersistStreamInit",IID_IPersistStreamInit},{"IPersistStorage",IID_IPersistStorage},{"IDataObject",IID_IDataObject},{"IPicture",IID_IPicture},{"IPictureDisp",IID_IPictureDisp},{"IViewObject",IID_IViewObject},{"IOleObject",IID_IOleObject},{"IServiceProvider",IID_IServiceProvider},{"IProvideClassInfo",IID_IProvideClassInfo}};
 for(auto& e:entries){IUnknown* result=nullptr;auto hr=object->QueryInterface(e.id,(void**)&result);out<<label<<" QI "<<e.name<<" 0x"<<std::hex<<(unsigned)hr<<std::dec<<'\n';if(result)result->Release();}
 for(auto name:{L"SetImage",L"SetImageBytes",L"SetImageFromStream",L"UserPictureFromStream",L"Picture",L"Texture",L"UserPicture"}){auto n=const_cast<wchar_t*>(name);DISPID id=DISPID_UNKNOWN;auto hr=object->GetIDsOfNames(IID_NULL,&n,1,LOCALE_USER_DEFAULT,&id);out<<label<<" member ";while(*name)out<<(char)*name++;out<<" hr=0x"<<std::hex<<(unsigned)hr<<std::dec<<" id="<<id<<'\n';}
}
int wmain(int argc,wchar_t** argv){
 if(argc!=2)return 2;
 auto root=std::filesystem::absolute(argv[1]);std::ofstream out(root/L"artifacts/interface_probe.txt");
 try{check(CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED),"COM");CLSID cls;check(CLSIDFromProgID(L"PowerPoint.Application",&cls),"CLSID");Value app;app.v.vt=VT_DISPATCH;check(CoCreateInstance(cls,nullptr,CLSCTX_LOCAL_SERVER,IID_IDispatch,(void**)&app.v.pdispVal),"App");
 auto pres=call(get(app.obj(),L"Presentations").obj(),L"Add",{Value(0L)});
 try{
 auto slide=call(get(pres.obj(),L"Slides").obj(),L"Add",{Value(1L),Value(12L)});auto shapes=get(slide.obj(),L"Shapes");auto shape=call(shapes.obj(),L"AddShape",{Value(1L),Value(10.),Value(10.),Value(100.),Value(100.)});auto fill=get(shape.obj(),L"Fill");probe(shape.obj(),"Shape",out);probe(fill.obj(),"FillFormat",out);
 std::ifstream input(root/L"artifacts/textures/texture_64_0.png",std::ios::binary);std::vector<char> bytes((std::istreambuf_iterator<char>(input)),{});if(bytes.empty())throw Error(E_FAIL,"Missing texture");
 Value array;array.v.vt=VT_ARRAY|VT_UI1;array.v.parray=SafeArrayCreateVector(VT_UI1,0,(ULONG)bytes.size());void* p;check(SafeArrayAccessData(array.v.parray,&p),"array");memcpy(p,bytes.data(),bytes.size());SafeArrayUnaccessData(array.v.parray);
 auto attempt=[&](const char* label,Value argument){try{call(fill.obj(),L"UserPicture",{argument});out<<label<<" returned success; fill="<<get(fill.obj(),L"Type").integer()<<'\n';}catch(const Error& e){out<<label<<" hr=0x"<<std::hex<<(unsigned)e.hr<<std::dec<<" "<<e.what()<<'\n';}};
 attempt("UserPicture(SAFEARRAY UI1)",array);
 IStream* stream=nullptr;check(CreateStreamOnHGlobal(nullptr,TRUE,&stream),"stream");ULONG written;check(stream->Write(bytes.data(),(ULONG)bytes.size(),&written),"stream write");LARGE_INTEGER zero{};stream->Seek(zero,STREAM_SEEK_SET,nullptr);Value sv;sv.v.vt=VT_UNKNOWN;sv.v.punkVal=stream;attempt("UserPicture(IStream)",sv);
 attempt("UserPicture(data URI)",Value(L"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6pVQAAAAASUVORK5CYII="));
 }catch(...){put(pres.obj(),L"Saved",Value(-1L));call(pres.obj(),L"Close");throw;}
 put(pres.obj(),L"Saved",Value(-1L));call(pres.obj(),L"Close");return 0;
 }catch(const Error& e){out<<e.what()<<" 0x"<<std::hex<<(unsigned)e.hr<<'\n';return 1;}catch(...){return 1;}
}
