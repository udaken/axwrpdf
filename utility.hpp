#include <string>
#include <vector>
#include <windows.h>
#include <assert.h>

static std::string w2string(LPCWSTR lpszStr, DWORD dwFlags = WC_NO_BEST_FIT_CHARS, LPCSTR  lpDefaultChar = nullptr)
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

static std::wstring a2wstring(LPCSTR lpszStr, DWORD dwFlags = 0)
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


static BOOL readAllBytes(LPCWSTR path, LPVOID data, size_t size)
{
	auto hFile = ::CreateFile(path, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);

	if (hFile == INVALID_HANDLE_VALUE)
	{
		return FALSE;
	}

	do
	{
		DWORD dwNumberOfBytesToRead = static_cast<DWORD>(min(UINT32_MAX, size));
		DWORD dwNumberOfBytesRead;

		if (!::ReadFile(hFile, data, dwNumberOfBytesToRead, &dwNumberOfBytesRead, nullptr) || dwNumberOfBytesRead != dwNumberOfBytesToRead)
		{
			::CloseHandle(hFile);
			return FALSE;
		}
		data = (reinterpret_cast<BYTE*>(data) + dwNumberOfBytesRead);
		size -= dwNumberOfBytesRead;
	} while (size > 0);

	::CloseHandle(hFile);
	return TRUE;
}

