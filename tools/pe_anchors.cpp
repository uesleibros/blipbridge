#include <windows.h>
#include <filesystem>
#include <fstream>
#include <vector>
#include <string>
#include <iostream>
#include <cstring>
int wmain(int argc,wchar_t** argv){
 if(argc<3)return 2;
 std::ifstream file(std::filesystem::path(argv[1]),std::ios::binary);std::vector<char> data((std::istreambuf_iterator<char>(file)),{});
 if(data.size()<sizeof(IMAGE_DOS_HEADER)){std::cerr<<"Cannot read PE file\n";return 1;}
 auto dos=(IMAGE_DOS_HEADER*)data.data();
 if(dos->e_magic!=IMAGE_DOS_SIGNATURE||dos->e_lfanew<0||size_t(dos->e_lfanew)+sizeof(IMAGE_NT_HEADERS64)>data.size())return 1;
 auto nt=(IMAGE_NT_HEADERS64*)(data.data()+dos->e_lfanew);if(nt->Signature!=IMAGE_NT_SIGNATURE||nt->FileHeader.Machine!=IMAGE_FILE_MACHINE_AMD64)return 1;
 auto sections=IMAGE_FIRST_SECTION(nt);if((char*)(sections+nt->FileHeader.NumberOfSections)>data.data()+data.size())return 1;
 auto ptr=[&](DWORD rva)->char*{for(int i=0;i<nt->FileHeader.NumberOfSections;i++){auto& s=sections[i];if(rva>=s.VirtualAddress&&rva-s.VirtualAddress<s.SizeOfRawData&&size_t(s.PointerToRawData)+s.SizeOfRawData<=data.size())return data.data()+s.PointerToRawData+(rva-s.VirtualAddress);}return nullptr;};
 auto directory=nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXCEPTION];auto pdata=(RUNTIME_FUNCTION*)ptr(directory.VirtualAddress);
 auto function=[&](DWORD rva){if(pdata)for(DWORD i=0;i<directory.Size/sizeof(RUNTIME_FUNCTION);i++)if(pdata[i].BeginAddress<=rva&&rva<pdata[i].EndAddress)return pdata[i].BeginAddress;return DWORD(0);};
 for(int arg=2;arg<argc;arg++){
  std::string needle;for(auto c:std::wstring(argv[arg]))needle.push_back((char)c);
  for(int i=0;i<nt->FileHeader.NumberOfSections;i++){
   auto& s=sections[i];if(size_t(s.PointerToRawData)+s.SizeOfRawData>data.size())continue;
   auto begin=data.data()+s.PointerToRawData;std::string_view view(begin,s.SizeOfRawData);size_t pos=0;
   while((pos=view.find(needle,pos))!=std::string_view::npos){DWORD target=s.VirtualAddress+(DWORD)pos++;std::cout<<needle<<" stringRVA=0x"<<std::hex<<target<<'\n';
    for(int j=0;j<nt->FileHeader.NumberOfSections;j++){
     auto& code=sections[j];if(!(code.Characteristics&IMAGE_SCN_MEM_EXECUTE)||size_t(code.PointerToRawData)+code.SizeOfRawData>data.size())continue;
     auto bytes=(unsigned char*)data.data()+code.PointerToRawData;
     for(DWORD k=0;k+7<=code.SizeOfRawData;k++)if((bytes[k]==0x48||bytes[k]==0x4c)&&bytes[k+1]==0x8d&&(bytes[k+2]&0xc7)==5){int32_t d;memcpy(&d,bytes+k+3,4);DWORD address=code.VirtualAddress+k;if(uint32_t(address+7+d)==target)std::cout<<"  LEA xrefRVA=0x"<<address<<" unwindRegionBeginRVA=0x"<<function(address)<<'\n';}
    }
   }
  }
 }
 return 0;
}
