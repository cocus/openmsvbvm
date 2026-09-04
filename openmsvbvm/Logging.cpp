#include "Logging.hpp"

#include <new>

namespace vbl
{

/*
 * Logger
 */

Logger & Logger::Get()
{
	static Logger s_logger;
	return s_logger;
}

Logger::Logger()
	: m_hFile(INVALID_HANDLE_VALUE)
{
	SetLevel(LOG_INFO);
	OpenFile();
}

Logger::~Logger()
{
	if (m_hFile != INVALID_HANDLE_VALUE)
	{
		CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
	}
}

void Logger::OpenFile()
{
	wchar_t dir[MAX_PATH];
	if (!GetTempPathW(MAX_PATH, dir))
	{
		return;
	}

	wchar_t path[MAX_PATH];
	wsprintfW(path, L"%lsopenmsvbvm.log", dir);

	// Append across runs: OPEN_ALWAYS + seek to end.
	m_hFile = CreateFileW(path, GENERIC_WRITE, FILE_SHARE_READ,
	                      NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (m_hFile != INVALID_HANDLE_VALUE)
	{
		SetFilePointer(m_hFile, 0, NULL, FILE_END);
	}
}

void Logger::SetLevel(LogLevel level)
{
	if (level >= LOG_TRACE && level < LOG_NONE)
	{
		m_level = level;
	}
}

bool Logger::IsEnabled(LogLevel level) const
{
	// m_level is a minimum-severity threshold: a message is shown when its
	// level is at least as severe (>=, not <=) as the configured level.
	return (level >= LOG_TRACE && level < LOG_NONE) && level >= m_level;
}

const wchar_t * Logger::LevelName(LogLevel level)
{
	static const wchar_t * s_names[LOG_NONE] = {
		L"TRACE", L"DEBUG", L"INFO", L"WARN", L"ERROR"
	};

	if (level >= LOG_TRACE && level < LOG_NONE)
	{
		return s_names[level];
	}
	return L"?????";
}

void Logger::Write(LogLevel level, const wchar_t * func, const void * obj, const wchar_t * text)
{
	if (!IsEnabled(level) || !text)
	{
		return;
	}

	if (!func)
	{
		func = L"";
	}

	const wchar_t * levelName = LevelName(level);

	// Heap-allocate the composed line to fit func/text exactly, rather than
	// wsprintfW-ing (no bounds checking at all) into a fixed stack array that
	// a long function name or message could overflow.
	size_t needed = wcslen(levelName) + wcslen(func) + wcslen(text) + 64;
	wchar_t * line = new (std::nothrow) wchar_t[needed];
	if (!line)
	{
		return;
	}

	if (obj)
	{
		_snwprintf_s(line, needed, _TRUNCATE, L"%-5ls %ls(%08lx): %ls\r\n",
		             levelName, func, (unsigned long)(size_t)obj, text);
	}
	else
	{
		_snwprintf_s(line, needed, _TRUNCATE, L"%-5ls %ls: %ls\r\n",
		             levelName, func, text);
	}

	if (m_hFile != INVALID_HANDLE_VALUE)
	{
		DWORD written = 0;
		WriteFile(m_hFile, line, (DWORD)(wcslen(line) * sizeof(wchar_t)), &written, NULL);
		FlushFileBuffers(m_hFile);
	}

	OutputDebugStringW(line);

	delete[] line;
}

/*
 * LogStream
 */

LogStream::LogStream(LogLevel level, const wchar_t * func, const void * obj)
	: m_level(level), m_func(func), m_obj(obj), m_enabled(Logger::Get().IsEnabled(level))
{
}

LogStream::~LogStream()
{
	if (m_enabled)
	{
		Logger::Get().Write(m_level, m_func, m_obj, m_stream.str().c_str());
	}
}

void LogStream::AppendGuid(REFGUID g)
{
	LPOLESTR psz = NULL;
	if (SUCCEEDED(StringFromCLSID(g, &psz)) && psz)
	{
		m_stream << psz;
		CoTaskMemFree(psz);
	}
	else
	{
		m_stream << L"(guid)";
	}
}

void LogStream::AppendHres(HRESULT hr)
{
	wchar_t buf[64];

	switch (hr)
	{
	case S_OK:            _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx (S_OK)", (unsigned long)hr); break;
	case E_FAIL:          _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx (E_FAIL)", (unsigned long)hr); break;
	case E_POINTER:       _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx (E_POINTER)", (unsigned long)hr); break;
	case E_UNEXPECTED:    _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx (E_UNEXPECTED)", (unsigned long)hr); break;
	case E_OUTOFMEMORY:   _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx (E_OUTOFMEMORY)", (unsigned long)hr); break;
	case E_NOINTERFACE:   _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx (E_NOINTERFACE)", (unsigned long)hr); break;
	case E_NOTIMPL:       _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx (E_NOTIMPL)", (unsigned long)hr); break;
	default:              _snwprintf_s(buf, 64, _TRUNCATE, L"0x%08lx", (unsigned long)hr); break;
	}

	m_stream << buf;
}

void LogStream::AppendBstr(BSTR b)
{
	if (!b)
	{
		m_stream << L"(null)";
	}
	else if (*b == L'\0')
	{
		m_stream << L"(empty)";
	}
	else
	{
		m_stream << b;
	}
}

void LogStream::AppendHex(unsigned long v, int width)
{
	wchar_t buf[32];
	_snwprintf_s(buf, 32, _TRUNCATE, L"%0*lx", width, v);
	m_stream << buf;
}

LogStream & LogStream::operator<<(const Guid & g)
{
	if (m_enabled)
	{
		AppendGuid(*g.value);
	}
	return *this;
}

LogStream & LogStream::operator<<(const Hres & h)
{
	if (m_enabled)
	{
		AppendHres(h.value);
	}
	return *this;
}

LogStream & LogStream::operator<<(const Bstr & b)
{
	if (m_enabled)
	{
		AppendBstr(b.value);
	}
	return *this;
}

LogStream & LogStream::operator<<(const Narrow & n)
{
	if (m_enabled)
	{
		if (!n.value)
		{
			m_stream << L"(null)";
		}
		else
		{
			int needed = MultiByteToWideChar(CP_ACP, 0, n.value, -1, NULL, 0);
			if (needed > 0)
			{
				wchar_t * wide = new (std::nothrow) wchar_t[needed];
				if (wide)
				{
					MultiByteToWideChar(CP_ACP, 0, n.value, -1, wide, needed);
					m_stream << wide;
					delete[] wide;
				}
			}
		}
	}
	return *this;
}

LogStream & LogStream::operator<<(const Var & v)
{
	if (m_enabled)
	{
		if (!v.value)
		{
			m_stream << L"(null)";
			return *this;
		}

		const VARIANT & var = *v.value;
		m_stream << L"vt=" << (unsigned short)var.vt << L" ";

		switch (var.vt & VT_TYPEMASK)
		{
		case VT_BSTR:  m_stream << L"bstr='"; AppendBstr(var.bstrVal); m_stream << L"'"; break;
		case VT_I2:    m_stream << L"i2="  << var.iVal; break;
		case VT_I4:    m_stream << L"i4="  << var.lVal; break;
		case VT_UI2:   m_stream << L"ui2=" << var.uiVal; break;
		case VT_UI4:   m_stream << L"ui4=" << var.ulVal; break;
		case VT_R4:    m_stream << L"r4="  << (double)var.fltVal; break;
		case VT_R8:    m_stream << L"r8="  << var.dblVal; break;
		case VT_BOOL:  m_stream << L"bool="<< (short)var.boolVal; break;
		default:       m_stream << L"(other)"; break;
		}
	}
	return *this;
}

LogStream & LogStream::operator<<(const Hex & h)
{
	if (m_enabled)
	{
		AppendHex(h.value, h.width);
	}
	return *this;
}

LogStream & LogStream::operator<<(const CY & cy)
{
	if (m_enabled)
	{
		m_stream << cy.int64;
	}
	return *this;
}

} // namespace vbl
