#include <string>
#include <vector>
#include <windows.h>
#include <assert.h>

inline std::string w2string(LPCWSTR lpszStr, DWORD dwFlags = WC_NO_BEST_FIT_CHARS, LPCSTR  lpDefaultChar = nullptr)
{
	size_t bufLen = wcslen(lpszStr) * 2 + 1;
	assert(bufLen <= INT32_MAX);
	std::vector<CHAR> buf(bufLen, '\0');
	int ret = ::WideCharToMultiByte(CP_ACP, dwFlags, lpszStr, -1, buf.data(), static_cast<int>(bufLen), lpDefaultChar, nullptr);

	if (ret == 0 && ::GetLastError() == ERROR_INSUFFICIENT_BUFFER)
	{
		buf.reserve(buf.capacity() + 256);
		::WideCharToMultiByte(CP_ACP, dwFlags, lpszStr, -1, buf.data(), static_cast<int>(bufLen), nullptr, nullptr);
	}

	return buf.data();
}

inline std::wstring a2wstring(LPCSTR lpszStr, DWORD dwFlags = 0)
{
	size_t bufLen = strlen(lpszStr) + 1;
	std::vector<WCHAR> buf(bufLen, L'\0');
	int ret = ::MultiByteToWideChar(CP_ACP, dwFlags, lpszStr, -1, buf.data(), static_cast<int>(bufLen));

	if (ret == 0 && ::GetLastError() == ERROR_INSUFFICIENT_BUFFER)
	{
		buf.reserve(buf.capacity() + 256);
		::MultiByteToWideChar(CP_ACP, dwFlags, lpszStr, -1, buf.data(), static_cast<int>(bufLen));
	}

	return buf.data();
}
