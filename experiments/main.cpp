#include <blipbridge/dispatch.hpp>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iomanip>
#include <iostream>
#include <numeric>
#include <psapi.h>
using namespace bb;
namespace fs = std::filesystem;
static double freq;

static long long ticks() {
    LARGE_INTEGER x;
    QueryPerformanceCounter(&x);
    return x.QuadPart;
}

static void bench(std::ofstream& f, const char* name, int n, const std::function<void(int)>& op) {
    std::vector<double> times;
    times.reserve(n);
    for (int i = 0; i < 10; i++) {
        op(i);
    }
    for (int i = 0; i < n; i++) {
        auto t = ticks();
        op(i);
        times.push_back((ticks() - t) * 1e6 / freq);
    }
    auto total = std::accumulate(times.begin(), times.end(), 0.0);
    std::sort(times.begin(), times.end());
    f << name << ',' << n << ',' << total / 1000 << ',' << total / n << ',' << times[n / 2] << ','
      << times.front() << ',' << times.back() << ',' << n * 1e6 / total << '\n';
    f.flush();
    std::cout << name << " n=" << n << " avg_us=" << total / n << std::endl;
}

static Value rect(IDispatch* shapes) {
    return call(shapes, L"AddShape", {Value(1L), Value(10.), Value(10.), Value(80.), Value(60.)});
}

static Value poly(IDispatch* shapes, double x = 100., double y = 100.) {
    auto b = call(shapes, L"BuildFreeform", {Value(0L), Value(x), Value(y)});
    for (auto p : std::vector<std::pair<double, double>>{{x + 40, y}, {x + 20, y + 40}, {x, y}}) {
        call(b.obj(), L"AddNodes", {Value(0L), Value(0L), Value(p.first), Value(p.second)});
    }
    return call(b.obj(), L"ConvertToShape");
}

static std::wstring snapshot(IDispatch* s) {
    std::wstring r;
    for (auto n : {L"Name",
                   L"Id",
                   L"Type",
                   L"Left",
                   L"Top",
                   L"Width",
                   L"Height",
                   L"Rotation",
                   L"ZOrderPosition"}) {
        r += std::wstring(n) + L"=" + get(s, n).str() + L";";
    }
    auto nodes = get(s, L"Nodes");
    r += L"Nodes=" + get(nodes.obj(), L"Count").str();
    return r;
}

static std::string narrow(const std::wstring& s) {
    std::string r;
    for (auto c : s) {
        r.push_back(c < 128 ? (char)c : '?');
    }
    return r;
}

static std::string typeName(ITypeInfo* info, const TYPEDESC& d) {
    if (d.vt == VT_PTR) {
        return typeName(info, *d.lptdesc) + "*";
    }
    if (d.vt == VT_SAFEARRAY) {
        return "SAFEARRAY(" + typeName(info, *d.lptdesc) + ")";
    }
    if (d.vt == VT_USERDEFINED) {
        ITypeInfo* t = nullptr;
        if (SUCCEEDED(info->GetRefTypeInfo(d.hreftype, &t))) {
            BSTR b = nullptr;
            t->GetDocumentation(MEMBERID_NIL, &b, nullptr, nullptr, nullptr);
            auto s = narrow(b ? b : L"?");
            SysFreeString(b);
            t->Release();
            return s;
        }
    }
    switch (d.vt) {
    case VT_BSTR:
        return "BSTR";
    case VT_VARIANT:
        return "VARIANT";
    case VT_DISPATCH:
        return "IDispatch";
    case VT_UNKNOWN:
        return "IUnknown";
    case VT_HRESULT:
        return "HRESULT";
    case VT_VOID:
        return "void";
    case VT_I4:
        return "I4";
    case VT_R4:
        return "R4";
    case VT_UI1:
        return "UI1";
    default:
        return "VT_" + std::to_string(d.vt);
    }
}

void dumpLibrary(ITypeLib* lib, std::ofstream& out) {
    TLIBATTR* la = nullptr;
    lib->GetLibAttr(&la);
    out << "Library version " << la->wMajorVerNum << '.' << la->wMinorVerNum
        << " syskind=" << la->syskind << '\n';
    lib->ReleaseTLibAttr(la);
    for (UINT i = 0; i < lib->GetTypeInfoCount(); i++) {
        ITypeInfo* t = nullptr;
        check(lib->GetTypeInfo(i, &t), "type");
        TYPEATTR* a = nullptr;
        t->GetTypeAttr(&a);
        BSTR b = nullptr;
        t->GetDocumentation(MEMBERID_NIL, &b, nullptr, nullptr, nullptr);
        auto name = narrow(b ? b : L"");
        SysFreeString(b);
        out << "\nTYPE " << name << " kind=" << a->typekind << " flags=0x" << std::hex
            << a->wTypeFlags << std::dec << " vtableBytes=" << a->cbSizeVft << '\n';
        for (UINT j = 0; j < a->cFuncs; j++) {
            FUNCDESC* d = nullptr;
            t->GetFuncDesc(j, &d);
            BSTR names[64]{};
            UINT count = 0;
            t->GetNames(d->memid, names, 64, &count);
            out << "  " << (count ? narrow(names[0]) : "?") << " dispid=" << d->memid
                << " invkind=" << d->invkind << " flags=0x" << std::hex << d->wFuncFlags << std::dec
                << " vtbl=" << d->oVft << " returns=" << typeName(t, d->elemdescFunc.tdesc) << " (";
            for (int k = 0; k < d->cParams; k++) {
                if (k) {
                    out << ", ";
                }
                out << typeName(t, d->lprgelemdescParam[k].tdesc)
                    << " flags=" << d->lprgelemdescParam[k].paramdesc.wParamFlags;
            }
            out << ")\n";
            for (UINT k = 0; k < count; k++) {
                SysFreeString(names[k]);
            }
            t->ReleaseFuncDesc(d);
        }
        t->ReleaseTypeAttr(a);
        t->Release();
    }
}

static void dumpObjectLibrary(IDispatch* o, const fs::path& file) {
    ITypeInfo* t = nullptr;
    check(o->GetTypeInfo(0, 0, &t), "GetTypeInfo");
    ITypeLib* l = nullptr;
    UINT i;
    check(t->GetContainingTypeLib(&l, &i), "GetContainingTypeLib");
    std::ofstream out(file);
    dumpLibrary(l, out);
    l->Release();
    t->Release();
}

int runExperiment(int argc, wchar_t** argv) {
    try {
        check(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED), "CoInitialize");
        LARGE_INTEGER q;
        QueryPerformanceFrequency(&q);
        freq = (double)q.QuadPart;
        fs::path root = fs::absolute(argc > 1 ? argv[1] : L".");
        fs::create_directories(root / L"artifacts");
        CLSID cls;
        check(CLSIDFromProgID(L"PowerPoint.Application", &cls), "PowerPoint CLSID");
        Value app;
        app.v.vt = VT_DISPATCH;
        check(CoCreateInstance(
                  cls, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&app.v.pdispVal),
              "PowerPoint activation");
        put(app.obj(), L"Visible", Value(-1L));
        auto presentations = get(app.obj(), L"Presentations");
        auto pres = call(presentations.obj(), L"Add", {Value(-1L)});
        auto slides = get(pres.obj(), L"Slides");
        auto slide = call(slides.obj(), L"Add", {Value(1L), Value(12L)});
        auto shapes = get(slide.obj(), L"Shapes");
        auto target = poly(shapes.obj());
        put(target.obj(), L"Name", Value(L"BB_original_freeform"));
        auto donor = rect(shapes.obj());
        auto donor2 = rect(shapes.obj());
        auto fill = get(target.obj(), L"Fill");
        bool inproc = GetModuleHandleW(L"POWERPNT.EXE") != nullptr;
        std::ofstream findings(
            root / (inproc ? L"artifacts/functional_inproc.txt" : L"artifacts/functional.txt"));
        findings << "PowerPoint " << narrow(get(app.obj(), L"Version").str())
                 << " inProcess=" << inproc << " PID=" << GetCurrentProcessId() << "\n";
        dumpObjectLibrary(app.obj(), root / L"artifacts/powerpoint_typelib.txt");
        dumpObjectLibrary(fill.obj(), root / L"artifacts/fill_typelib.txt");
        auto tex = [&](int s, int v) {
            return (root / L"artifacts/textures" /
                    (L"texture_" + std::to_wstring(s) + L"_" + std::to_wstring(v) + L".png"))
                .wstring();
        };
        auto user = [&](IDispatch* f, const std::wstring& p) {
            call(f, L"UserPicture", {Value(p.c_str())});
        };
        user(get(donor.obj(), L"Fill").obj(), tex(64, 0));
        user(get(donor2.obj(), L"Fill").obj(), tex(64, 1));
        if (argc > 2 && std::wstring(argv[2]) == L"trace") {
            std::cout << "TRACE READY: waiting 15 seconds" << std::endl;
            Sleep(15000);
            user(fill.obj(),
                 (root / L"artifacts/textures/BB_TRACE_UNIQUE_TEXTURE_01.png").wstring());
            std::cout << "TRACE CALL COMPLETE" << std::endl;
            put(pres.obj(), L"Saved", Value(-1L));
            call(pres.obj(), L"Close");
            return 0;
        }
        auto before = snapshot(target.obj());
        user(fill.obj(), tex(64, 0));
        findings << "UserPicture geometry identity preserved=" << (before == snapshot(target.obj()))
                 << '\n';
        call(donor.obj(), L"PickUp");
        call(target.obj(), L"Apply");
        findings << "PickUp/Apply geometry identity preserved="
                 << (before == snapshot(target.obj()))
                 << " fillType=" << get(fill.obj(), L"Type").integer() << '\n';
        auto nodes = get(target.obj(), L"Nodes");
        call(nodes.obj(), L"SetPosition", {Value(2L), Value(151.), Value(100.)});
        findings << "Freeform Nodes.SetPosition succeeded\n";
        // Deliberately test whether Apply also transfers non-fill style.
        auto dl = get(donor.obj(), L"Line");
        auto tl = get(target.obj(), L"Line");
        put(dl.obj(), L"Weight", Value(7.));
        put(tl.obj(), L"Weight", Value(1.));
        call(donor.obj(), L"PickUp");
        call(target.obj(), L"Apply");
        findings << "Apply target line weight after donor weight 7="
                 << get(tl.obj(), L"Weight").number() << '\n';
        findings.flush();
        std::ofstream csv(root /
                          (inproc ? L"artifacts/baseline_inproc.csv" : L"artifacts/baseline.csv"));
        csv << "operation,count,total_ms,average_us,median_us,min_us,max_us,ops_per_sec\n";
        for (int n : {100, 1000, 10000}) {
            bench(csv, "UserPicture_same_64", n, [&](int) { user(fill.obj(), tex(64, 0)); });
            bench(csv, "UserPicture_alternating_64", n, [&](int i) {
                user(fill.obj(), tex(64, i % 2));
            });
            call(donor.obj(), L"PickUp");
            bench(csv, "Apply_prePicked", n, [&](int) { call(target.obj(), L"Apply"); });
            bench(csv, "PickUp_Apply_alternating", n, [&](int i) {
                call(i % 2 ? donor.obj() : donor2.obj(), L"PickUp");
                call(target.obj(), L"Apply");
            });
        }
        for (int s : {32, 128, 256}) {
            bench(csv, ("UserPicture_same_" + std::to_string(s)).c_str(), 1000, [&](int) {
                user(fill.obj(), tex(s, 0));
            });
        }
        std::vector<Value> created;
        bench(csv, "Shape_creation", 1000, [&](int) { created.push_back(rect(shapes.obj())); });
        bench(csv, "Shape_deletion", 1000, [&](int) {
            call(created.back().obj(), L"Delete");
            created.pop_back();
        });
        bench(csv, "Duplicate_and_delete", 1000, [&](int) {
            auto d = call(target.obj(), L"Duplicate");
            call(d.obj(), L"Delete");
        });
        bench(csv, "Node_SetPosition", 1000, [&](int i) {
            call(nodes.obj(), L"SetPosition", {Value(2L), Value(140. + i % 2), Value(100.)});
        });
        bench(csv, "Visibility_change", 1000, [&](int i) {
            put(target.obj(), L"Visible", Value(i % 2 ? -1L : 0L));
        });
        std::vector<Value> pool, fills, poolNodes;
        for (int i = 0; i < 130; i++) {
            pool.push_back(poly(shapes.obj(), (i % 13) * 50., (i / 13) * 45.));
            fills.push_back(get(pool.back().obj(), L"Fill"));
            poolNodes.push_back(get(pool.back().obj(), L"Nodes"));
        }
        for (bool pickup : {false, true}) {
            bench(csv, pickup ? "Frame130_PickUp_Apply" : "Frame130_UserPicture", 100, [&](int f) {
                for (int i = 0; i < 130; i++) {
                    if (pickup) {
                        call((i + f) % 2 ? donor.obj() : donor2.obj(), L"PickUp");
                        call(pool[i].obj(), L"Apply");
                    } else {
                        user(fills[i].obj(), tex(64, (i + f) % 2));
                    }
                    call(poolNodes[i].obj(),
                         L"SetPosition",
                         {Value(2L), Value((i % 13) * 50. + 40. + f % 2), Value((i / 13) * 45.)});
                    put(pool[i].obj(), L"Visible", Value((i + f) % 7 ? -1L : 0L));
                }
            });
        }
        // Different filenames with identical bytes, plus duplication, for saved-resource
        // comparison.
        user(fill.obj(),
             (root / L"artifacts/textures/identical_bytes_other_filename.png").wstring());
        auto dupe = call(target.obj(), L"Duplicate");
        call(pres.obj(),
             L"SaveAs",
             {Value((root / L"artifacts/baseline.pptx").c_str()), Value(24L)});
        call(slide.obj(),
             L"Export",
             {Value((root / L"artifacts/baseline.png").c_str()),
              Value(L"PNG"),
              Value(1280L),
              Value(720L)});
        call(pres.obj(), L"Close");
        auto reopened = call(
            presentations.obj(),
            L"Open",
            {Value((root / L"artifacts/baseline.pptx").c_str()), Value(0L), Value(0L), Value(0L)});
        auto rs = get(reopened.obj(), L"Slides");
        auto rslide = item(rs.obj(), 1);
        auto rsh = get(rslide.obj(), L"Shapes");
        auto original = invoke(rsh.obj(),
                               L"Item",
                               DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                               {Value(L"BB_original_freeform")});
        findings << "Reopened original type=" << get(original.obj(), L"Type").integer()
                 << " fill=" << get(get(original.obj(), L"Fill").obj(), L"Type").integer() << '\n';
        call(reopened.obj(), L"Close");
        std::cout << "Completed baseline and save/reopen checks" << std::endl;
    } catch (const Error& e) {
        std::cerr << e.what() << " HRESULT=0x" << std::hex << (unsigned)e.hr << std::endl;
        return 1;
    } catch (const std::exception& e) {
        std::cerr << e.what() << std::endl;
        return 1;
    }
    return 0;
}
#ifndef BB_DLL
int wmain(int argc, wchar_t** argv) {
    return runExperiment(argc, argv);
}
#endif
