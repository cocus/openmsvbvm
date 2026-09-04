#pragma once

#include <windows.h>
#include <sstream>

#include <objbase.h> /* StringFromCLSID / CoTaskMemFree */
#include <oleauto.h> /* BSTR / VARIANT */

/*
 * C++03-compatible logging layer (deliberately kept strict C++03 so it can be
 * backported to older toolchains): no nullptr/override/auto/constexpr/variadic
 * templates/NSDMI. Uses 0/NULL, plain member init, and classic do-nothing
 * copy-ctor/assign declarations to make classes non-copyable.
 *
 * Usage is a stream, not a format string:
 *
 *   LOG(LOG_DEBUG) << L"caption='" << vbl::Bstr(caption) << L"' hr=" << vbl::Hres(hr);
 *   LOG_OBJ(LOG_WARN, this) << L"hwnd=" << vbl::Hex((unsigned long)hwnd);
 *
 * Plain types (ints, pointers, wide/narrow string literals, ...) stream via
 * std::wostringstream's own operator<<; the vbl:: wrapper types below exist
 * only for the handful of things that need special-cased formatting.
 */

/*
 * Log levels at global scope (this project is namespace-free, so LOG(...)'s
 * level argument must resolve unqualified from any .cpp) -- the order also
 * seeds Logger's default "TRACE off, everything else on" policy via a
 * numeric comparison.
 */
enum LogLevel
{
    LOG_TRACE = 0,
    LOG_DEBUG = 1,
    LOG_INFO = 2,
    LOG_WARN = 3,
    LOG_ERROR = 4,
    LOG_NONE = 5
};

namespace vbl
{

/*
 * Process-wide logger writing to BOTH %TEMP%\openmsvbvm.log (append mode) and
 * OutputDebugStringW, tagged by level and (optionally) the owning object.
 * Singleton reached through Get(); levels can be toggled at runtime.
 */
class Logger
{
public:
    static Logger& Get();

    void SetLevel(LogLevel level);
    bool IsEnabled(LogLevel level) const;

    void Write(LogLevel level, const wchar_t* func, const void* obj, const wchar_t* text);

private:
    Logger();
    ~Logger();
    Logger(const Logger&); // non-copyable (declared, not defined)
    Logger& operator=(const Logger&);

    void OpenFile();
    static const wchar_t* LevelName(LogLevel level);

    HANDLE m_hFile;
    LogLevel m_level; // set in the ctor via SetLevel(); no in-class initializer (C++03)
};

/*
 * Small value wrappers for the things that need special-cased formatting
 * when streamed into a LogStream: a GUID (via StringFromCLSID), an HRESULT
 * (best-effort symbolic name), a BSTR (null/empty guarded) and an ANSI C
 * string (properly widened via MultiByteToWideChar, not just zero-extended).
 * Everything else -- ints, pointers, wide string literals, narrow string
 * literals, doubles, bools -- streams directly via std::wostringstream's own
 * operator<< overloads, no wrapper needed.
 */
struct Guid
{
    explicit Guid(REFGUID guid) : value(&guid)
    {
    }
    const GUID* value;
};

struct Hres
{
    explicit Hres(HRESULT hr) : value(hr)
    {
    }
    HRESULT value;
};

struct Bstr
{
    explicit Bstr(BSTR b) : value(b)
    {
    }
    BSTR value;
};

struct Narrow
{
    explicit Narrow(const char* s) : value(s)
    {
    }
    const char* value;
};

struct Var
{
    explicit Var(const VARIANT& v) : value(&v)
    {
    }
    const VARIANT* value;
};

/*
 * Zero-padded hex, matching the old "%.8x" / "%08lx" style used throughout
 * the codebase for pointers, flags and error codes. Overloaded on the
 * argument's width (32/16/8-bit) so the default padding matches it (8/4/2
 * digits); pass an explicit width to override.
 *
 * `int`/`unsigned int` get their own overloads even though they're the same
 * width as `long`/`unsigned long` on this platform: `int` and `long` (and
 * their unsigned counterparts) are still distinct types, so plain `int`/
 * `unsigned int` arguments (UINT, a `char` promoted to `int`, ...) would
 * otherwise be an equally-ranked conversion to all of unsigned
 * char/short/long and fail to compile as ambiguous. The pointer overload
 * similarly covers any pointer type (LPVOID, struct pointers, ...) via the
 * standard pointer-to-const-void* conversion, without needing a cast at the
 * call site.
 */
struct Hex
{
    explicit Hex(unsigned long v, int w = 8) : value(v), width(w)
    {
    }
    explicit Hex(long v, int w = 8) : value((unsigned long)v), width(w)
    {
    }
    explicit Hex(unsigned int v, int w = 8) : value(v), width(w)
    {
    }
    explicit Hex(int v, int w = 8) : value((unsigned long)v), width(w)
    {
    }
    explicit Hex(unsigned short v, int w = 4) : value(v), width(w)
    {
    }
    explicit Hex(unsigned char v, int w = 2) : value(v), width(w)
    {
    }
    explicit Hex(char v, int w = 2) : value(v), width(w)
    {
    }
    explicit Hex(const void* v, int w = 8) : value((unsigned long)v), width(w)
    {
    }

    unsigned long value;
    int width;
};

/*
 * RAII log-line builder returned by LOG()/LOG_OBJ(). Streaming into it
 * (operator<<) appends to an internal buffer; when the temporary is
 * destroyed (end of the full expression) the composed line is handed to
 * Logger::Write. When the level is disabled every << is a no-op, so nothing
 * is ever formatted for a suppressed log line.
 */
class LogStream
{
public:
    LogStream(LogLevel level, const wchar_t* func, const void* obj);
    ~LogStream();

    template <typename T>
    LogStream& operator<<(const T& value)
    {
        if (m_enabled)
        {
            m_stream << value;
        }
        return *this;
    }

    LogStream& operator<<(const Guid& g);
    LogStream& operator<<(const Hres& h);
    LogStream& operator<<(const Bstr& b);
    LogStream& operator<<(const Narrow& n);
    LogStream& operator<<(const Var& v);
    LogStream& operator<<(const Hex& h);
    LogStream& operator<<(const CY& cy); // no built-in ostream inserter for this COM struct

private:
    LogStream(const LogStream&); // non-copyable (declared, not defined)
    LogStream& operator=(const LogStream&);

    void AppendGuid(REFGUID g);
    void AppendHres(HRESULT hr);
    void AppendBstr(BSTR b);
    void AppendHex(unsigned long v, int width);

    LogLevel m_level;
    const wchar_t* m_func;
    const void* m_obj;
    bool m_enabled;
    std::wostringstream m_stream;
};

} // namespace vbl

/*
 * Public macros. Each constructs exactly one vbl::LogStream temporary tagged
 * with __FUNCTION__; all real logic lives in C++. The temporary lives for the
 * rest of the full expression, so every << chained onto it runs before the
 * line is written out when it goes out of scope.
 */
#define LOG(LEVEL) vbl::LogStream((LEVEL), L"" __FUNCTION__, (const void*)0)
#define LOG_OBJ(LEVEL, SELF) vbl::LogStream((LEVEL), L"" __FUNCTION__, (const void*)(SELF))
