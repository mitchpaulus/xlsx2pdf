#define NOMINMAX
#include <windows.h>
#include <comdef.h>
#include <filesystem>
#include <iostream>
#include <string>
#include <string_view>
#include <vector>
#include <thread>
#include <chrono>
#include <sstream>
#include <stdexcept>
#include <iomanip>
#include <algorithm>
#include <limits>
#include <comip.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

namespace {

std::string Narrow(const std::wstring &value) {
    if (value.empty()) {
        return {};
    }
    int required = ::WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0) {
        return {};
    }
    std::string result(static_cast<size_t>(required), '\0');
    ::WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), result.data(), required, nullptr, nullptr);
    return result;
}

std::wstring GetEnvironmentValue(const wchar_t *name) {
    DWORD size = ::GetEnvironmentVariableW(name, nullptr, 0);
    if (size == 0) {
        return {};
    }
    std::wstring buffer(size, L'\0');
    DWORD written = ::GetEnvironmentVariableW(name, buffer.data(), size);
    if (written == 0 || written >= size) {
        return {};
    }
    buffer.resize(written);
    return buffer;
}

_variant_t VariantMissing() {
    _variant_t missing;
    missing.vt = VT_ERROR;
    missing.scode = DISP_E_PARAMNOTFOUND;
    return missing;
}

IDispatch *RequireDispatch(const _variant_t &variant, const char *context) {
    if (variant.vt != VT_DISPATCH || variant.pdispVal == nullptr) {
        std::ostringstream stream;
        stream << "Expected COM dispatch for " << context;
        throw std::runtime_error(stream.str());
    }
    return variant.pdispVal;
}

long RequireLong(const _variant_t &variant, const char *context) {
    if (variant.vt == VT_I4) {
        return variant.lVal;
    }
    if (variant.vt == VT_I2) {
        return variant.iVal;
    }
    _variant_t converted(variant);
    HRESULT hr = VariantChangeType(static_cast<VARIANT *>(&converted), static_cast<const VARIANT *>(&converted), 0, VT_I4);
    if (SUCCEEDED(hr)) {
        return converted.lVal;
    }
    std::ostringstream stream;
    stream << "Expected integer result for " << context;
    throw std::runtime_error(stream.str());
}

_variant_t Invoke(IDispatch *dispatch, WORD flags, LPCOLESTR name, std::initializer_list<_variant_t> args = {}) {
    if (!dispatch) {
        throw std::runtime_error("COM dispatch pointer is null");
    }

    LPOLESTR names[] = { const_cast<LPOLESTR>(name) };
    DISPID dispid = 0;
    HRESULT hr = dispatch->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    std::vector<VARIANT> reversedArgs(args.size());
    for (auto &variant : reversedArgs) {
        VariantInit(&variant);
    }

    size_t index = 0;
    for (auto it = args.end(); it != args.begin();) {
        --it;
        const VARIANT &source = *it;
        hr = VariantCopy(&reversedArgs[index], const_cast<VARIANT *>(&source));
        if (FAILED(hr)) {
            for (size_t cleanup = 0; cleanup <= index && cleanup < reversedArgs.size(); ++cleanup) {
                VariantClear(&reversedArgs[cleanup]);
            }
            _com_issue_error(hr);
        }
        ++index;
    }

    DISPPARAMS params{};
    params.cArgs = static_cast<UINT>(reversedArgs.size());
    params.rgvarg = reversedArgs.empty() ? nullptr : reversedArgs.data();
    params.cNamedArgs = 0;

    DISPID namedArg = DISPID_PROPERTYPUT;
    if (flags & DISPATCH_PROPERTYPUT) {
        params.cNamedArgs = 1;
        params.rgdispidNamedArgs = &namedArg;
    }

    _variant_t result;
    hr = dispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, flags, &params, &result, nullptr, nullptr);

    for (auto &variant : reversedArgs) {
        VariantClear(&variant);
    }

    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    return result;
}

class ComInitializer {
public:
    ComInitializer() {
        HRESULT hr = ::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (FAILED(hr)) {
            if (hr == RPC_E_CHANGED_MODE) {
                throw std::runtime_error("COM has already been initialized with a different threading model.");
            }
            _com_issue_error(hr);
        }
        initialized_ = true;
    }

    ComInitializer(const ComInitializer &) = delete;
    ComInitializer &operator=(const ComInitializer &) = delete;

    ~ComInitializer() {
        if (initialized_) {
            ::CoUninitialize();
        }
    }

private:
    bool initialized_ = false;
};

// Excel's COM server frequently returns RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER
// while it is busy (initialization, add-in loads, printer driver round-trips, modal
// prompts elsewhere on the desktop). Registering an IMessageFilter on the STA tells
// the COM runtime to wait briefly and retry instead of failing the call.
class RetryMessageFilter : public IMessageFilter {
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv) override {
        if (!ppv) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IMessageFilter) {
            *ppv = static_cast<IMessageFilter *>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&refCount_));
    }

    ULONG STDMETHODCALLTYPE Release() override {
        LONG count = InterlockedDecrement(&refCount_);
        return static_cast<ULONG>(count);
    }

    DWORD STDMETHODCALLTYPE HandleInComingCall(DWORD, HTASK, DWORD, LPINTERFACEINFO) override {
        return SERVERCALL_ISHANDLED;
    }

    DWORD STDMETHODCALLTYPE RetryRejectedCall(HTASK, DWORD dwTickCount, DWORD dwRejectType) override {
        if (dwRejectType != SERVERCALL_RETRYLATER && dwRejectType != SERVERCALL_REJECTED) {
            return static_cast<DWORD>(-1);
        }
        // Give Excel up to ~30 seconds of cumulative retry time, polling every 250 ms.
        if (dwTickCount >= 30000) {
            return static_cast<DWORD>(-1);
        }
        return 250;
    }

    DWORD STDMETHODCALLTYPE MessagePending(HTASK, DWORD, DWORD) override {
        return PENDINGMSG_WAITDEFPROCESS;
    }

private:
    LONG refCount_ = 1;
};

class MessageFilterScope {
public:
    MessageFilterScope() {
        filter_ = new RetryMessageFilter();
        IMessageFilter *previous = nullptr;
        HRESULT hr = ::CoRegisterMessageFilter(filter_, &previous);
        if (FAILED(hr)) {
            filter_->Release();
            filter_ = nullptr;
            _com_issue_error(hr);
        }
        previous_ = previous;
        registered_ = true;
    }

    MessageFilterScope(const MessageFilterScope &) = delete;
    MessageFilterScope &operator=(const MessageFilterScope &) = delete;

    ~MessageFilterScope() {
        if (registered_) {
            IMessageFilter *current = nullptr;
            ::CoRegisterMessageFilter(previous_, &current);
            if (current) {
                current->Release();
            }
        }
        if (filter_) {
            filter_->Release();
        }
        if (previous_) {
            previous_->Release();
        }
    }

private:
    RetryMessageFilter *filter_ = nullptr;
    IMessageFilter *previous_ = nullptr;
    bool registered_ = false;
};

using DispatchPtr = IDispatchPtr;

DispatchPtr RequireDispatchPtr(_variant_t &variant, const char *context) {
    DispatchPtr result(RequireDispatch(variant, context));
    variant.Clear();
    return result;
}

class UniqueHandle {
public:
    UniqueHandle() = default;
    explicit UniqueHandle(HANDLE handle) : handle_(handle) {}
    UniqueHandle(const UniqueHandle &) = delete;
    UniqueHandle &operator=(const UniqueHandle &) = delete;
    UniqueHandle(UniqueHandle &&other) noexcept : handle_(other.handle_) {
        other.handle_ = nullptr;
    }
    UniqueHandle &operator=(UniqueHandle &&other) noexcept {
        if (this != &other) {
            Reset(other.Release());
        }
        return *this;
    }
    ~UniqueHandle() {
        Reset();
    }

    HANDLE Get() const {
        return handle_;
    }

    HANDLE Release() {
        HANDLE handle = handle_;
        handle_ = nullptr;
        return handle;
    }

    void Reset(HANDLE handle = nullptr) {
        if (handle_ && handle_ != handle) {
            ::CloseHandle(handle_);
        }
        handle_ = handle;
    }

    explicit operator bool() const {
        return handle_ != nullptr;
    }

private:
    HANDLE handle_ = nullptr;
};

DWORD RequireProcessId(HWND windowHandle, const char *context) {
    DWORD processId = 0;
    if (!windowHandle) {
        std::ostringstream stream;
        stream << "Window handle unavailable for " << context;
        throw std::runtime_error(stream.str());
    }
    if (::GetWindowThreadProcessId(windowHandle, &processId) == 0 || processId == 0) {
        std::ostringstream stream;
        stream << "Failed to query process id for " << context;
        throw std::runtime_error(stream.str());
    }
    return processId;
}

std::chrono::milliseconds ClampToDword(std::chrono::milliseconds duration) {
    constexpr DWORD maxWait = std::numeric_limits<DWORD>::max();
    auto count = duration.count();
    if (count < 0) {
        return std::chrono::milliseconds(0);
    }
    if (static_cast<unsigned long long>(count) > maxWait) {
        return std::chrono::milliseconds(maxWait);
    }
    return duration;
}

bool WaitForExit(UniqueHandle &handle, std::chrono::milliseconds timeout) {
    if (!handle) {
        return true;
    }
    DWORD result = ::WaitForSingleObject(handle.Get(), static_cast<DWORD>(ClampToDword(timeout).count()));
    if (result == WAIT_OBJECT_0) {
        return true;
    }
    if (result == WAIT_FAILED) {
        std::ostringstream stream;
        stream << "WaitForSingleObject failed with error " << ::GetLastError();
        throw std::runtime_error(stream.str());
    }
    return false;
}

UniqueHandle OpenProcessHandle(DWORD processId, DWORD desiredAccess) {
    HANDLE handle = ::OpenProcess(desiredAccess, FALSE, processId);
    return UniqueHandle(handle);
}

UniqueHandle AcquireExcelProcessHandle(const DispatchPtr &excelApplication, DWORD &processId) {
    constexpr int kMaxAttempts = 30;
    for (int attempt = 0; attempt < kMaxAttempts; ++attempt) {
        _variant_t hwndVariant;
        try {
            hwndVariant = Invoke(excelApplication.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"Hwnd");
        } catch (...) {
            std::this_thread::sleep_for(std::chrono::milliseconds(200));
            continue;
        }

        long hwndValue = RequireLong(hwndVariant, "Application.Hwnd");
        HWND hwnd = reinterpret_cast<HWND>(static_cast<intptr_t>(hwndValue));
        if (!hwnd) {
            std::this_thread::sleep_for(std::chrono::milliseconds(200));
            continue;
        }

        try {
            processId = RequireProcessId(hwnd, "Excel application");
        } catch (...) {
            std::this_thread::sleep_for(std::chrono::milliseconds(200));
            continue;
        }

        UniqueHandle handle = OpenProcessHandle(processId, SYNCHRONIZE);
        if (handle) {
            return handle;
        }

        // If we cannot open the process yet, wait and retry.
        std::this_thread::sleep_for(std::chrono::milliseconds(200));
    }

    throw std::runtime_error("Unable to acquire handle for Excel process.");
}

void EnsureProcessTermination(DWORD processId, UniqueHandle &processHandle) {
    if (processId == 0) {
        return;
    }

    constexpr auto waitBeforeTerminate = std::chrono::seconds(15);
    if (processHandle && WaitForExit(processHandle, waitBeforeTerminate)) {
        processHandle.Reset();
        return;
    }

    processHandle.Reset();

    UniqueHandle terminateHandle = OpenProcessHandle(processId, PROCESS_TERMINATE | SYNCHRONIZE);
    if (!terminateHandle) {
        std::ostringstream stream;
        stream << "Excel process (PID " << processId << ") did not exit and could not be opened for termination.";
        throw std::runtime_error(stream.str());
    }

    if (!::TerminateProcess(terminateHandle.Get(), 1)) {
        std::ostringstream stream;
        stream << "Excel process (PID " << processId << ") did not exit and termination failed (error "
               << ::GetLastError() << ").";
        throw std::runtime_error(stream.str());
    }

    constexpr auto waitAfterTerminate = std::chrono::seconds(5);
    (void)WaitForExit(terminateHandle, waitAfterTerminate);
}

} // namespace

int wmain(int argc, wchar_t *argv[]) {
    auto printUsage = [](std::wostream &stream) {
        stream << L"Usage: xlsx2pdf [options] <input-path> [worksheet-name]\n"
               << L"  --landscape, -l    Export the page in landscape orientation.\n"
               << L"  --portrait,  -p    Export the page in portrait orientation (default).\n"
               << L"  --fit-to-page, -f  Scale the worksheet to fit on a single page.\n"
               << L"  -h, --help         Show this help and exit.\n"
               << L"Converts the specified Excel worksheet to PDF and saves it to %TMP%\\xlsx.pdf.\n"
               << L"Provide an optional worksheet name to export a specific sheet; defaults to the first sheet.\n";
    };

    std::vector<std::wstring_view> positional;
    bool landscape = false;
    bool fitToPage = false;

    for (int i = 1; i < argc; ++i) {
        std::wstring_view arg(argv[i]);
        if (arg == L"-h" || arg == L"--help" || arg == L"/?") {
            printUsage(std::wcout);
            return 0;
        }
        if (arg == L"--landscape" || arg == L"-l") {
            landscape = true;
        } else if (arg == L"--portrait" || arg == L"-p") {
            landscape = false;
        } else if (arg == L"--fit-to-page" || arg == L"-f") {
            fitToPage = true;
        } else if (!arg.empty() && arg.front() == L'-') {
            std::wcerr << L"Unknown option: " << arg << L"\n";
            printUsage(std::wcerr);
            return 1;
        } else {
            positional.push_back(arg);
        }
    }

    if (positional.empty() || positional.size() > 2) {
        printUsage(std::wcerr);
        return 1;
    }

    std::filesystem::path inputPath(positional[0]);
    std::error_code ec;
    inputPath = std::filesystem::absolute(inputPath, ec);
    if (ec) {
        std::cerr << "Failed to resolve input path.\n";
        return 1;
    }

    if (!std::filesystem::exists(inputPath)) {
        std::wcerr << L"Input file not found: " << inputPath.native() << L"\n";
        return 1;
    }

    if (!std::filesystem::is_regular_file(inputPath)) {
        std::wcerr << L"Input path is not a file: " << inputPath.native() << L"\n";
        return 1;
    }

    std::wstring worksheetName;
    if (positional.size() == 2) {
        worksheetName = std::wstring(positional[1]);
    }

    std::wstring tmpDirectory = GetEnvironmentValue(L"TMP");
    if (tmpDirectory.empty()) {
        std::cerr << "TMP environment variable is not set.\n";
        return 1;
    }

    std::filesystem::path outputPath = std::filesystem::path(tmpDirectory) / L"xlsx.pdf";

    try {
        ComInitializer com;
        MessageFilterScope messageFilter;

        CLSID clsid;
        HRESULT hr = ::CLSIDFromProgID(L"Excel.Application", &clsid);
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }

        IDispatch *excelRaw = nullptr;
        hr = ::CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, reinterpret_cast<void **>(&excelRaw));
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }

        DispatchPtr excel;
        excel.Attach(excelRaw);
        excelRaw = nullptr;

        DWORD excelProcessId = 0;
        UniqueHandle excelProcessHandle;
        try {
            excelProcessHandle = AcquireExcelProcessHandle(excel, excelProcessId);
        } catch (...) {
            try {
                Invoke(excel.GetInterfacePtr(), DISPATCH_METHOD, L"Quit");
            } catch (...) {
                // Best effort; ignore failures during cleanup.
            }
            excel = nullptr;
            EnsureProcessTermination(excelProcessId, excelProcessHandle);
            throw;
        }

        bool quitCalled = false;
        bool workbookOpened = false;

        DispatchPtr workbooks;
        DispatchPtr workbook;
        DispatchPtr worksheets;
        DispatchPtr worksheet;

        _variant_t workbooksVariant;
        _variant_t workbookVariant;
        _variant_t worksheetsVariant;

        try {
            Invoke(excel.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"Visible", { _variant_t(VARIANT_FALSE, VT_BOOL) });
            Invoke(excel.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"DisplayAlerts", { _variant_t(VARIANT_FALSE, VT_BOOL) });

            std::wcerr << L"Exporting " << inputPath.native() << L"\n";

            workbooksVariant = Invoke(excel.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"Workbooks");
            workbooks = RequireDispatchPtr(workbooksVariant, "Application.Workbooks");

            workbookVariant = Invoke(workbooks.GetInterfacePtr(), DISPATCH_METHOD, L"Open", { _variant_t(inputPath.c_str()) });
            workbookOpened = true;

            workbook = RequireDispatchPtr(workbookVariant, "Workbooks.Open result");

            worksheetsVariant = Invoke(workbook.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"Worksheets");
            worksheets = RequireDispatchPtr(worksheetsVariant, "Workbook.Worksheets");

            constexpr int kMaxWorksheetPollAttempts = 30;
            int pollAttempts = 0;
            long sheetCount = 0;
            while (true) {
                _variant_t countVariant = Invoke(worksheets.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"Count");
                sheetCount = RequireLong(countVariant, "Worksheets.Count");
                if (sheetCount > 0) {
                    break;
                }
                if (++pollAttempts >= kMaxWorksheetPollAttempts) {
                    throw std::runtime_error("Timed out waiting for workbook worksheets to load.");
                }
                std::this_thread::sleep_for(std::chrono::seconds(1));
            }

            if (worksheetName.empty()) {
                _variant_t firstWorksheetVariant = Invoke(worksheets.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"Item", { _variant_t(1L) });
                worksheet = RequireDispatchPtr(firstWorksheetVariant, "Worksheets.Item");
            } else {
                bool found = false;
                std::vector<std::wstring> availableSheets;
                for (long index = 1; index <= sheetCount; ++index) {
                    _variant_t sheetVariant = Invoke(worksheets.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"Item", { _variant_t(index) });
                    DispatchPtr sheetDispatch = RequireDispatchPtr(sheetVariant, "Worksheets.Item");
                    _variant_t nameVariant = Invoke(sheetDispatch.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"Name");
                    std::wstring sheetName = nameVariant.bstrVal ? std::wstring(nameVariant.bstrVal) : std::wstring();
                    availableSheets.push_back(sheetName);
                    if (sheetName == worksheetName) {
                        worksheet = sheetDispatch;
                        found = true;
                        break;
                    }
                }

                if (!found) {
                    std::wostringstream stream;
                    stream << L"Worksheet '" << worksheetName << L"' not found in the workbook. Possible sheets include:";
                    for (const auto &sheet : availableSheets) {
                        stream << L" " << sheet;
                    }
                    throw std::runtime_error(Narrow(stream.str()));
                }
            }

            if (!worksheet) {
                throw std::runtime_error("Worksheet dispatch is not available.");
            }

            if (landscape || fitToPage) {
                _variant_t pageSetupVariant = Invoke(worksheet.GetInterfacePtr(), DISPATCH_PROPERTYGET, L"PageSetup");
                DispatchPtr pageSetup = RequireDispatchPtr(pageSetupVariant, "Worksheet.PageSetup");

                if (landscape) {
                    // xlLandscape = 2, xlPortrait = 1
                    Invoke(pageSetup.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"Orientation", { _variant_t(2L) });
                }

                if (fitToPage) {
                    // Zoom must be disabled before FitToPagesWide/Tall take effect.
                    Invoke(pageSetup.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"Zoom", { _variant_t(VARIANT_FALSE, VT_BOOL) });
                    Invoke(pageSetup.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"FitToPagesWide", { _variant_t(1L) });
                    Invoke(pageSetup.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"FitToPagesTall", { _variant_t(1L) });
                }
            }

            Invoke(worksheet.GetInterfacePtr(), DISPATCH_METHOD, L"ExportAsFixedFormat", {
                _variant_t(0L),
                _variant_t(outputPath.c_str()),
                _variant_t(0L),
                _variant_t(VARIANT_TRUE, VT_BOOL),
                _variant_t(VARIANT_FALSE, VT_BOOL),
                VariantMissing(),
                VariantMissing(),
                _variant_t(VARIANT_FALSE, VT_BOOL),
                VariantMissing()
            });

            std::wcerr << L"PDF export successful: " << outputPath.native() << L"\n";

            worksheet = nullptr;
            worksheets = nullptr;

            if (workbookOpened && workbook) {
                Invoke(workbook.GetInterfacePtr(), DISPATCH_METHOD, L"Close", { _variant_t(VARIANT_FALSE, VT_BOOL) });
                workbookOpened = false;
            }

            workbook = nullptr;
            workbooks = nullptr;

            Invoke(excel.GetInterfacePtr(), DISPATCH_METHOD, L"Quit");
            quitCalled = true;
        } catch (...) {
            worksheet = nullptr;
            worksheets = nullptr;
            if (workbookOpened && workbook) {
                try {
                    Invoke(workbook.GetInterfacePtr(), DISPATCH_METHOD, L"Close", { _variant_t(VARIANT_FALSE, VT_BOOL) });
                } catch (...) {
                    // swallow cleanup failure
                }
            }
            if (!quitCalled && excel) {
                try {
                    Invoke(excel.GetInterfacePtr(), DISPATCH_METHOD, L"Quit");
                } catch (...) {
                    // swallow cleanup failure
                }
            }
            workbook = nullptr;
            workbooks = nullptr;
            excel = nullptr;
            EnsureProcessTermination(excelProcessId, excelProcessHandle);
            throw;
        }

        excel = nullptr;
        EnsureProcessTermination(excelProcessId, excelProcessHandle);

        return 0;
    } catch (const _com_error &e) {
        std::wcerr << L"An error occurred: " << e.ErrorMessage() << L" (HRESULT: 0x"
                   << std::hex << std::uppercase << e.Error() << L")\n";
        return 1;
    } catch (const std::exception &e) {
        std::cerr << "An error occurred: " << e.what() << "\n";
        return 1;
    }
}
