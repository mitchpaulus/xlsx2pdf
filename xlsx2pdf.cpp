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
#include <optional>
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

std::wstring Widen(const std::string &value) {
    if (value.empty()) {
        return {};
    }
    int required = ::MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0);
    if (required <= 0) {
        return {};
    }
    std::wstring result(static_cast<size_t>(required), L'\0');
    ::MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), result.data(), required);
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

// Resolves an input path to absolute and verifies it is an existing regular
// file. Returns an empty string on success, otherwise a description of the
// problem. The path is rewritten to its absolute form on success.
std::wstring ValidateInputPath(std::filesystem::path &path) {
    std::error_code ec;
    path = std::filesystem::absolute(path, ec);
    if (ec) {
        return L"Failed to resolve input path.";
    }
    if (!std::filesystem::exists(path)) {
        return L"Input file not found: " + path.wstring();
    }
    if (!std::filesystem::is_regular_file(path)) {
        return L"Input path is not a file: " + path.wstring();
    }
    return {};
}

// Resolves an output path to absolute and verifies its parent directory exists.
// Returns an empty string on success, otherwise a description of the problem.
// The path is rewritten to its absolute form on success.
std::wstring ValidateOutputPath(std::filesystem::path &path) {
    std::error_code ec;
    path = std::filesystem::absolute(path, ec);
    if (ec) {
        return L"Failed to resolve output path.";
    }
    std::filesystem::path dir = path.parent_path();
    if (!dir.empty() && !std::filesystem::exists(dir)) {
        return L"Output directory does not exist: " + dir.wstring();
    }
    return {};
}

// Opens the workbook at inputPath inside the supplied (already-running) Excel
// instance, exports the requested worksheet (or the first sheet when
// worksheetName is empty) to outputPath, and closes the workbook again. The
// Excel instance itself is left running so a single instance can service many
// conversions. Throws on failure.
void ExportWorksheet(IDispatch *excel,
                     const std::filesystem::path &inputPath,
                     const std::wstring &worksheetName,
                     bool landscape,
                     bool fitToPage,
                     const std::filesystem::path &outputPath) {
    DispatchPtr workbooks;
    DispatchPtr workbook;
    DispatchPtr worksheets;
    DispatchPtr worksheet;
    bool workbookOpened = false;

    _variant_t workbooksVariant;
    _variant_t workbookVariant;
    _variant_t worksheetsVariant;

    try {
        workbooksVariant = Invoke(excel, DISPATCH_PROPERTYGET, L"Workbooks");
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

        worksheet = nullptr;
        worksheets = nullptr;

        Invoke(workbook.GetInterfacePtr(), DISPATCH_METHOD, L"Close", { _variant_t(VARIANT_FALSE, VT_BOOL) });
        workbookOpened = false;

        workbook = nullptr;
        workbooks = nullptr;
    } catch (...) {
        worksheet = nullptr;
        worksheets = nullptr;
        if (workbookOpened && workbook) {
            try {
                Invoke(workbook.GetInterfacePtr(), DISPATCH_METHOD, L"Close", { _variant_t(VARIANT_FALSE, VT_BOOL) });
            } catch (...) {
                // swallow cleanup failure; the workbook reference is released below
            }
        }
        workbook = nullptr;
        workbooks = nullptr;
        throw;
    }
}

// Owns a dedicated, hidden Excel automation instance for the lifetime of the
// object. The constructor launches Excel, takes a handle on its process, and
// suppresses UI; the destructor quits Excel and waits for (or terminates) the
// process. Reusing one instance across many conversions avoids paying Excel's
// startup and shutdown cost per file.
class ExcelInstance {
public:
    ExcelInstance() {
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
        excel_.Attach(excelRaw);

        try {
            excelProcessHandle_ = AcquireExcelProcessHandle(excel_, excelProcessId_);
            Invoke(excel_.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"Visible", { _variant_t(VARIANT_FALSE, VT_BOOL) });
            Invoke(excel_.GetInterfacePtr(), DISPATCH_PROPERTYPUT, L"DisplayAlerts", { _variant_t(VARIANT_FALSE, VT_BOOL) });
        } catch (...) {
            Shutdown();
            throw;
        }
    }

    ExcelInstance(const ExcelInstance &) = delete;
    ExcelInstance &operator=(const ExcelInstance &) = delete;

    ~ExcelInstance() {
        Shutdown();
    }

    IDispatch *Get() const {
        return excel_.GetInterfacePtr();
    }

private:
    void Shutdown() noexcept {
        if (excel_) {
            try {
                Invoke(excel_.GetInterfacePtr(), DISPATCH_METHOD, L"Quit");
            } catch (...) {
                // best effort; fall through to process termination
            }
            excel_ = nullptr;
        }
        try {
            EnsureProcessTermination(excelProcessId_, excelProcessHandle_);
        } catch (...) {
            // best effort during teardown; never throw from here
        }
        excelProcessId_ = 0;
    }

    DispatchPtr excel_;
    DWORD excelProcessId_ = 0;
    UniqueHandle excelProcessHandle_;
};

} // namespace

// Reads tab-separated rows from stdin (UTF-8) and converts each one using a
// single shared Excel instance. Each row is "<input>\t<output>[\t<worksheet>]".
// Per-row failures are reported but do not stop the batch; the return value is
// non-zero if any row failed.
int RunBatch(bool landscape, bool fitToPage, bool skipExists) {
    try {
        ComInitializer com;
        MessageFilterScope messageFilter;
        // Launched lazily so a fully-skipped run (e.g. resuming with
        // --skip-exists) never pays Excel's startup cost.
        std::optional<ExcelInstance> excel;

        // Buffer every non-blank row up front so the progress counter knows the
        // total. A batch list is small relative to the conversions themselves.
        std::vector<std::string> rows;
        {
            std::string rawLine;
            bool firstLine = true;
            while (std::getline(std::cin, rawLine)) {
                // stdin is opened in text mode, but normalize defensively in case
                // a lone CR survives (e.g. piped from a tool that emits \r\n).
                if (!rawLine.empty() && rawLine.back() == '\r') {
                    rawLine.pop_back();
                }
                // Strip a UTF-8 BOM if the file carries one on its first line.
                if (firstLine && rawLine.size() >= 3 &&
                    static_cast<unsigned char>(rawLine[0]) == 0xEF &&
                    static_cast<unsigned char>(rawLine[1]) == 0xBB &&
                    static_cast<unsigned char>(rawLine[2]) == 0xBF) {
                    rawLine.erase(0, 3);
                }
                firstLine = false;

                if (!rawLine.empty()) {
                    rows.push_back(rawLine);
                }
            }
        }

        const size_t total = rows.size();
        size_t succeeded = 0;
        size_t failed = 0;
        size_t skipped = 0;

        for (size_t index = 0; index < total; ++index) {
            std::wostringstream prefixStream;
            prefixStream << L"(" << (index + 1) << L"/" << total << L") ";
            const std::wstring prefix = prefixStream.str();

            std::vector<std::string> fields;
            size_t start = 0;
            while (true) {
                size_t tab = rows[index].find('\t', start);
                if (tab == std::string::npos) {
                    fields.push_back(rows[index].substr(start));
                    break;
                }
                fields.push_back(rows[index].substr(start, tab - start));
                start = tab + 1;
            }

            if (fields.size() < 2 || fields[0].empty() || fields[1].empty()) {
                std::wcerr << prefix << L"FAIL\texpected <input-path>\\t<output-pdf> columns.\n";
                ++failed;
                continue;
            }

            std::filesystem::path inputPath = Widen(fields[0]);
            std::filesystem::path outputPath = Widen(fields[1]);
            std::wstring worksheetName = fields.size() >= 3 ? Widen(fields[2]) : std::wstring();

            if (std::wstring error = ValidateInputPath(inputPath); !error.empty()) {
                std::wcerr << prefix << L"FAIL\t" << error << L"\n";
                ++failed;
                continue;
            }
            if (std::wstring error = ValidateOutputPath(outputPath); !error.empty()) {
                std::wcerr << prefix << L"FAIL\t" << error << L"\n";
                ++failed;
                continue;
            }

            if (skipExists && std::filesystem::exists(outputPath)) {
                std::wcerr << prefix << L"SKIP\t" << outputPath.native() << L" already exists\n";
                ++skipped;
                continue;
            }

            // Launch Excel on first real work. A failure here is an
            // environment problem, so let it abort the whole batch.
            if (!excel) {
                excel.emplace();
            }

            std::wcerr << prefix << L"Converting " << inputPath.native() << L" -> " << outputPath.native() << L"\n";

            try {
                ExportWorksheet(excel->Get(), inputPath, worksheetName, landscape, fitToPage, outputPath);
                std::wcerr << prefix << L"OK\n";
                ++succeeded;
            } catch (const _com_error &e) {
                std::wcerr << prefix << L"FAIL\t" << inputPath.native() << L": " << e.ErrorMessage()
                           << L" (HRESULT: 0x" << std::hex << std::uppercase << e.Error() << std::dec << L")\n";
                ++failed;
            } catch (const std::exception &e) {
                std::wcerr << prefix << L"FAIL\t" << inputPath.native() << L": " << Widen(e.what()) << L"\n";
                ++failed;
            }
        }

        std::wcerr << L"Batch complete: " << succeeded << L" succeeded, " << failed
                   << L" failed, " << skipped << L" skipped, " << total << L" total.\n";
        return failed == 0 ? 0 : 1;
    } catch (const _com_error &e) {
        std::wcerr << L"An error occurred: " << e.ErrorMessage() << L" (HRESULT: 0x"
                   << std::hex << std::uppercase << e.Error() << L")\n";
        return 1;
    } catch (const std::exception &e) {
        std::cerr << "An error occurred: " << e.what() << "\n";
        return 1;
    }
}

int wmain(int argc, wchar_t *argv[]) {
    auto printUsage = [](std::wostream &stream) {
        stream << L"Usage: xlsx2pdf [options] <input-path> [worksheet-name]\n"
               << L"       xlsx2pdf batch [options]\n"
               << L"\n"
               << L"Options (apply to every exported sheet):\n"
               << L"  --landscape, -l       Export the page in landscape orientation.\n"
               << L"  --portrait,  -p       Export the page in portrait orientation (default).\n"
               << L"  --fit-to-page, -f     Scale the worksheet to fit on a single page.\n"
               << L"  --skip-exists, -s     Skip conversion when the output PDF already exists.\n"
               << L"  --output, -o <path>   Write the PDF to <path> (default: %TMP%\\xlsx.pdf).\n"
               << L"                        Not valid with the batch subcommand.\n"
               << L"  -h, --help            Show this help and exit.\n"
               << L"\n"
               << L"Single mode converts one worksheet to PDF. Provide an optional worksheet\n"
               << L"name to export a specific sheet; defaults to the first sheet.\n"
               << L"\n"
               << L"Batch mode reads tab-separated rows from stdin and converts them all using a\n"
               << L"single Excel instance. Columns: <input-path>\\t<output-pdf>[\\t<worksheet-name>].\n"
               << L"Blank lines are skipped. Each row's result is reported to stderr and the exit\n"
               << L"code is non-zero if any row failed.\n";
    };

    bool batchMode = false;
    int argStart = 1;
    if (argc > 1 && std::wstring_view(argv[1]) == L"batch") {
        batchMode = true;
        argStart = 2;
    }

    std::vector<std::wstring_view> positional;
    bool landscape = false;
    bool fitToPage = false;
    bool skipExists = false;
    std::wstring_view outputArg;

    for (int i = argStart; i < argc; ++i) {
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
        } else if (arg == L"--skip-exists" || arg == L"-s") {
            skipExists = true;
        } else if (arg == L"--output" || arg == L"-o") {
            if (i + 1 >= argc) {
                std::wcerr << L"Option " << arg << L" requires a path argument.\n";
                printUsage(std::wcerr);
                return 1;
            }
            outputArg = std::wstring_view(argv[++i]);
        } else if (!arg.empty() && arg.front() == L'-') {
            std::wcerr << L"Unknown option: " << arg << L"\n";
            printUsage(std::wcerr);
            return 1;
        } else {
            positional.push_back(arg);
        }
    }

    if (batchMode) {
        if (!positional.empty()) {
            std::wcerr << L"The batch subcommand does not take positional arguments; rows are read from stdin.\n";
            printUsage(std::wcerr);
            return 1;
        }
        if (!outputArg.empty()) {
            std::wcerr << L"The --output option is not valid with the batch subcommand; use the second column instead.\n";
            printUsage(std::wcerr);
            return 1;
        }
        return RunBatch(landscape, fitToPage, skipExists);
    }

    if (positional.empty() || positional.size() > 2) {
        printUsage(std::wcerr);
        return 1;
    }

    std::filesystem::path inputPath(positional[0]);
    if (std::wstring error = ValidateInputPath(inputPath); !error.empty()) {
        std::wcerr << error << L"\n";
        return 1;
    }

    std::wstring worksheetName;
    if (positional.size() == 2) {
        worksheetName = std::wstring(positional[1]);
    }

    std::filesystem::path outputPath;
    if (outputArg.empty()) {
        std::wstring tmpDirectory = GetEnvironmentValue(L"TMP");
        if (tmpDirectory.empty()) {
            std::cerr << "TMP environment variable is not set.\n";
            return 1;
        }
        outputPath = std::filesystem::path(tmpDirectory) / L"xlsx.pdf";
    } else {
        outputPath = std::filesystem::path(outputArg);
        if (std::wstring error = ValidateOutputPath(outputPath); !error.empty()) {
            std::wcerr << error << L"\n";
            return 1;
        }
    }

    try {
        if (skipExists && std::filesystem::exists(outputPath)) {
            std::wcerr << L"Skipping; output already exists: " << outputPath.native() << L"\n";
            return 0;
        }

        ComInitializer com;
        MessageFilterScope messageFilter;
        ExcelInstance excel;

        std::wcerr << L"Exporting " << inputPath.native() << L"\n";
        ExportWorksheet(excel.Get(), inputPath, worksheetName, landscape, fitToPage, outputPath);
        std::wcerr << L"PDF export successful: " << outputPath.native() << L"\n";

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
