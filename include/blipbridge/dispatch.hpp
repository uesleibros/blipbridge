#pragma once

#include <algorithm>
// MinGW requires Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>
#include <stdexcept>
#include <string>
#include <vector>

namespace bb {
/** Internal exception carrying an HRESULT; translate at every COM/C ABI boundary. */
struct Error : std::runtime_error {
    HRESULT hr;

    Error(HRESULT status, const std::string& message) : std::runtime_error(message), hr(status) {}
};

inline void check(HRESULT status, const char* context) {
    if (FAILED(status)) {
        throw Error(status, context);
    }
}

/**
 * Owning VARIANT with deep copy and noexcept transfer semantics.
 * Dispatch construction AddRefs; obj() returns a borrowed pointer valid for this
 * Value's lifetime. The public v member permits calls to native Automation APIs.
 */
struct Value {
    VARIANT v;

    Value() {
        VariantInit(&v);
    }

    explicit Value(long value) : Value() {
        v.vt = VT_I4;
        v.lVal = value;
    }

    explicit Value(double value) : Value() {
        v.vt = VT_R8;
        v.dblVal = value;
    }

    explicit Value(const wchar_t* value) : Value() {
        v.vt = VT_BSTR;
        v.bstrVal = SysAllocString(value);
        if (value && !v.bstrVal) {
            throw std::bad_alloc();
        }
    }

    explicit Value(IDispatch* object) : Value() {
        v.vt = VT_DISPATCH;
        v.pdispVal = object;
        if (object) {
            object->AddRef();
        }
    }

    Value(const Value& other) : Value() {
        // MinGW's Automation declaration lacks const on the source parameter.
        check(VariantCopy(&v, const_cast<VARIANT*>(&other.v)), "VariantCopy");
    }

    Value(Value&& other) noexcept : v(other.v) {
        VariantInit(&other.v);
    }

    Value& operator=(Value other) {
        std::swap(v, other.v);
        return *this;
    }

    ~Value() {
        VariantClear(&v);
    }

    IDispatch* obj() const {
        if (v.vt != VT_DISPATCH || !v.pdispVal) {
            throw Error(E_NOINTERFACE, "Expected dispatch");
        }
        return v.pdispVal;
    }

    long integer() const {
        Value converted;
        check(VariantChangeType(&converted.v, const_cast<VARIANT*>(&v), 0, VT_I4), "integer");
        return converted.v.lVal;
    }

    double number() const {
        Value converted;
        check(VariantChangeType(&converted.v, const_cast<VARIANT*>(&v), 0, VT_R8), "number");
        return converted.v.dblVal;
    }

    std::wstring str() const {
        Value converted;
        check(VariantChangeType(&converted.v, const_cast<VARIANT*>(&v), 0, VT_BSTR), "string");
        return converted.v.bstrVal
                   ? std::wstring(converted.v.bstrVal, SysStringLen(converted.v.bstrVal))
                   : L"";
    }
};

/** Owns the strings returned in EXCEPINFO, including on C++ allocation failure. */
struct ScopedExceptionInfo {
    EXCEPINFO value{};

    ScopedExceptionInfo() = default;
    ScopedExceptionInfo(const ScopedExceptionInfo&) = delete;
    ScopedExceptionInfo& operator=(const ScopedExceptionInfo&) = delete;

    ~ScopedExceptionInfo() {
        SysFreeString(value.bstrSource);
        SysFreeString(value.bstrDescription);
        SysFreeString(value.bstrHelpFile);
    }
};

/**
 * Calls a public Automation member on the current STA.
 * Argument Values own storage; raw variants below only borrow it during Invoke.
 * Name resolution remains per call, preserving the measured fallback behavior.
 */
inline Value
invoke(IDispatch* object, const wchar_t* name, WORD flags, std::vector<Value> arguments = {}) {
    if (!object) {
        throw Error(E_POINTER, "Missing dispatch object");
    }
    DISPID id;
    auto mutableName = const_cast<wchar_t*>(name);
    check(object->GetIDsOfNames(IID_NULL, &mutableName, 1, LOCALE_USER_DEFAULT, &id),
          "GetIDsOfNames");

    std::reverse(arguments.begin(), arguments.end());
    std::vector<VARIANT> rawArguments;
    rawArguments.reserve(arguments.size());
    for (auto& argument : arguments) {
        rawArguments.push_back(argument.v);
    }

    DISPID propertyPut = DISPID_PROPERTYPUT;
    DISPPARAMS parameters{rawArguments.data(), nullptr, static_cast<UINT>(rawArguments.size()), 0};
    if (flags & DISPATCH_PROPERTYPUT) {
        parameters.rgdispidNamedArgs = &propertyPut;
        parameters.cNamedArgs = 1;
    }

    Value result;
    ScopedExceptionInfo exception;
    UINT badArgument = 0;
    const HRESULT status = object->Invoke(id,
                                          IID_NULL,
                                          LOCALE_USER_DEFAULT,
                                          flags,
                                          &parameters,
                                          &result.v,
                                          &exception.value,
                                          &badArgument);
    // Avoid building diagnostic strings on successful hot-path calls.
    if (FAILED(status)) {
        std::string message = "Invoke ";
        for (const wchar_t* character = name; *character; ++character) {
            message.push_back(static_cast<char>(*character));
        }
        if (exception.value.bstrDescription) {
            message += " : ";
            for (auto character : std::wstring(exception.value.bstrDescription)) {
                message.push_back(character < 128 ? static_cast<char>(character) : '?');
            }
        }
        throw Error(status, message);
    }
    return result;
}

inline Value get(IDispatch* object, const wchar_t* name) {
    return invoke(object, name, DISPATCH_PROPERTYGET);
}

inline Value call(IDispatch* object, const wchar_t* name, std::vector<Value> arguments = {}) {
    return invoke(object, name, DISPATCH_METHOD, std::move(arguments));
}

inline void put(IDispatch* object, const wchar_t* name, Value value) {
    invoke(object, name, DISPATCH_PROPERTYPUT, {std::move(value)});
}

inline Value item(IDispatch* object, long index) {
    return invoke(object, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, {Value(index)});
}
} // namespace bb
