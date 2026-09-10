#include <blipbridge/dispatch.hpp>
#include <filesystem>
#include <fstream>
#include <numeric>
#include <psapi.h>
using namespace bb;
extern HRESULT memoryFillExperiment(IDispatch*, SAFEARRAY*, const std::wstring&);

static size_t privateBytes() {
    PROCESS_MEMORY_COUNTERS_EX p{};
    GetProcessMemoryInfo(GetCurrentProcess(), (PROCESS_MEMORY_COUNTERS*)&p, sizeof p);
    return p.PrivateUsage;
}

static long long qpc() {
    LARGE_INTEGER q;
    QueryPerformanceCounter(&q);
    return q.QuadPart;
}

template <class F>
static void timed(std::ofstream& out, const char* label, int n, F op) {
    LARGE_INTEGER f;
    QueryPerformanceFrequency(&f);
    std::vector<double> times;
    times.reserve(n);
    auto mem = privateBytes();
    for (int i = 0; i < n; i++) {
        auto start = qpc();
        op(i);
        times.push_back((qpc() - start) * 1e6 / f.QuadPart);
    }
    auto total = std::accumulate(times.begin(), times.end(), 0.);
    std::sort(times.begin(), times.end());
    out << label << ',' << n << ',' << total / 1000 << ',' << total / n << ',' << times[n / 2]
        << ',' << times.front() << ',' << times.back() << ',' << n * 1e6 / total << ',' << mem
        << ',' << privateBytes() << '\n';
    out.flush();
}

HRESULT stressExperiment(IDispatch* app, const std::wstring& root) {
    try {
        std::ofstream log(std::filesystem::path(root) / L"artifacts/stress.csv");
        log << "operation,count,total_ms,average_us,median_us,min_us,max_us,ops_per_sec,private_"
               "before,private_after\n";
        auto presentations = get(app, L"Presentations");
        auto pres = call(presentations.obj(), L"Add", {Value(0L)});
        try {
            auto slide = call(get(pres.obj(), L"Slides").obj(), L"Add", {Value(1L), Value(12L)});
            auto shapes = get(slide.obj(), L"Shapes");
            auto rect = [&]() {
                return call(shapes.obj(),
                            L"AddShape",
                            {Value(1L), Value(10.), Value(10.), Value(80.), Value(60.)});
            };
            auto donor = rect(), second = rect(), target = rect();
            std::vector<Value> donors{donor, second};
            std::vector<Value> images;
            for (int i = 0; i < 2; i++) {
                std::ifstream input(std::filesystem::path(root) / L"artifacts/textures" /
                                        (L"texture_64_" + std::to_wstring(i) + L".png"),
                                    std::ios::binary);
                std::vector<char> bytes((std::istreambuf_iterator<char>(input)), {});
                if (bytes.empty()) {
                    throw Error(E_FAIL, "Missing input");
                }
                Value array;
                array.v.vt = VT_ARRAY | VT_UI1;
                array.v.parray = SafeArrayCreateVector(VT_UI1, 0, (ULONG)bytes.size());
                if (!array.v.parray) {
                    throw std::bad_alloc();
                }
                void* p;
                check(SafeArrayAccessData(array.v.parray, &p), "array");
                memcpy(p, bytes.data(), bytes.size());
                SafeArrayUnaccessData(array.v.parray);
                images.push_back(std::move(array));
                check(memoryFillExperiment(donors[i].obj(), images[i].v.parray, root), "RAM donor");
            }
            timed(log, "MemoryAdapter_including_hooks_logging", 100, [&](int i) {
                check(memoryFillExperiment(target.obj(), images[i % 2].v.parray, root),
                      "RAM repeated");
            });
            images.clear();
            timed(log, "CachedDonor_100000_alternating", 100000, [&](int i) {
                call(donors[i % 2].obj(), L"PickUp");
                call(target.obj(), L"Apply");
            });
            std::vector<Value> pool;
            Value indices;
            indices.v.vt = VT_ARRAY | VT_VARIANT;
            indices.v.parray = SafeArrayCreateVector(VT_VARIANT, 0, 130);
            if (!indices.v.parray) {
                throw std::bad_alloc();
            }
            for (LONG i = 0; i < 130; i++) {
                pool.push_back(rect());
                auto name = L"BB_pool_" + std::to_wstring(i);
                put(pool.back().obj(), L"Name", Value(name.c_str()));
                Value v(name.c_str());
                check(SafeArrayPutElement(indices.v.parray, &i, &v.v), "range array");
            }
            auto range =
                invoke(shapes.obj(), L"Range", DISPATCH_METHOD | DISPATCH_PROPERTYGET, {indices});
            timed(log, "CachedDonor_130_individual_fills", 20, [&](int i) {
                call(donors[i % 2].obj(), L"PickUp");
                for (auto& s : pool) {
                    call(s.obj(), L"Apply");
                }
            });
            timed(log, "CachedDonor_130_ShapeRange_fills", 100, [&](int i) {
                call(donors[i % 2].obj(), L"PickUp");
                call(range.obj(), L"Apply");
            });
            call(pres.obj(),
                 L"SaveAs",
                 {Value((std::filesystem::path(root) / L"artifacts/stress.pptx").c_str()),
                  Value(24L)});
        } catch (...) {
            put(pres.obj(), L"Saved", Value(-1L));
            call(pres.obj(), L"Close");
            throw;
        }
        call(pres.obj(), L"Close");
        log << "# private bytes after document close=" << privateBytes() << '\n';
        return S_OK;
    } catch (const Error& e) {
        OutputDebugStringA(e.what());
        return e.hr;
    } catch (...) {
        return E_UNEXPECTED;
    }
}
