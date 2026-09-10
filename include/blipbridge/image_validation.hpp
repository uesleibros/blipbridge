#pragma once
#include <cstddef>
#include <windows.h>

namespace bb {
HRESULT validateImage(const BYTE* data, size_t size);
}
