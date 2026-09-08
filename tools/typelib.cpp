#include <blipbridge/dispatch.hpp>
#include <filesystem>
#include <fstream>
#include <iostream>
extern void dumpLibrary(ITypeLib*,std::ofstream&);
int wmain(int argc,wchar_t** argv){
 if(argc!=3){std::cerr<<"bb_typelib module-or-tlb output.txt\n";return 2;}
 HRESULT init=CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED);if(FAILED(init))return 1;
 ITypeLib* lib=nullptr;auto hr=LoadTypeLibEx(argv[1],REGKIND_NONE,&lib);
 if(FAILED(hr)){std::cerr<<"LoadTypeLibEx 0x"<<std::hex<<(unsigned)hr<<'\n';CoUninitialize();return 1;}
 try{std::ofstream out{std::filesystem::path(argv[2])};dumpLibrary(lib,out);}catch(...){lib->Release();CoUninitialize();return 1;}
 lib->Release();CoUninitialize();return 0;
}
