// axwinpdf.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
//

#include "stdafx.h"
#include <cstdlib>
#include <utility>
#include <vector>
#include <string>
#include <algorithm>
#include <memory>
#include <thread>
#include <ctime>

#include <windows.h>
#include <Shlwapi.h>
#include <shcore.h>
#include <assert.h>
#include <wincodec.h>
#include <comdef.h>

#include <windows.storage.h>
#include <windows.storage.streams.h>
#include <windows.data.pdf.h>
#include <windows.data.pdf.interop.h>

#include <wrl/client.h>
#include <wrl/async.h>
#include <wrl/event.h>
#include <wrl/wrappers/corewrappers.h> 

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "runtimeobject.lib")

#include "utility.hpp"

using namespace ABI::Windows::Foundation;
using namespace ABI::Windows::Data::Pdf;
using namespace ABI::Windows::Storage;
using namespace ABI::Windows::Storage::Streams;
using namespace Microsoft::WRL;
using namespace Microsoft::WRL::Wrappers;

#define CHECK_HR(...) \
    do { \
        HRESULT _hr = __VA_ARGS__; \
        if (FAILED(_hr)) { \
			_com_error err(_hr); \
            OutputDebugString(err.ErrorMessage()); \
            _com_issue_error(_hr); \
        } \
    } while (0)

#define REFCOUNT_DEBUG(async)  OutputDebugString((L"<<" L#async L">>ref count: " + std::to_wstring((async->AddRef(), async->Release())) + L"\n").c_str())


static HMODULE g_hModule = nullptr;
IPdfPage *_p = nullptr;
bool getAccurateSize = false;

// エラーコード
#define SPI_SUCCESS			0		// 正常終了
#define SPI_NOT_IMPLEMENT	(-1)	// その機能はインプリメントされていない
#define SPI_USER_CANCEL		1		// コールバック関数が非0を返したので展開を中止した
#define SPI_UNKNOWN_FORMAT	2		// 未知のフォーマット
#define SPI_DATA_BROKEN		3		// データが壊れている
#define SPI_NO_MEMORY		4		// メモリーが確保出来ない
#define SPI_MEMORY_ERR		5		// メモリーエラー（Lock出来ない、等）
#define SPI_FILE_READ_ERR	6		// ファイルリードエラー
#define SPI_RESERVED_ERR	7		// （予約）
#define SPI_INTERNAL_ERR	8		// 内部エラー

#ifdef _WIN64
typedef __time64_t		susie_time_t;
#else
typedef __time32_t		susie_time_t;
#endif

#include <pshpack1.h>
// ファイル情報構造体
struct fileInfo
{
	unsigned char	method[8];		// 圧縮法の種類
	ULONG_PTR		position;		// ファイル上での位置
	ULONG_PTR		compsize;		// 圧縮されたサイズ
	ULONG_PTR		filesize;		// 元のファイルサイズ
	susie_time_t	timestamp;		// ファイルの更新日時
	char			path[200];		// 相対パス
	char			filename[200];	// ファイルネーム
	unsigned long	crc;			// CRC
#ifdef _WIN64
									// 64bit版の構造体サイズは444bytesですが、実際のサイズは
									// アラインメントにより448bytesになります。環境によりdummyが必要です。
	char        dummy[4];
#endif
};
#include <poppack.h>

struct fileInfoW
{
	unsigned char	method[8];		// 圧縮法の種類
	ULONG_PTR		position;		// ファイル上での位置
	ULONG_PTR		compsize;		// 圧縮されたサイズ
	ULONG_PTR		filesize;		// 元のファイルサイズ
	susie_time_t	timestamp;		// ファイルの更新日時
	WCHAR			path[200];		// 相対パス
	WCHAR			filename[200];	// ファイルネーム
	unsigned long	crc;			// CRC
#ifdef _WIN64
	char        dummy[4];
#endif
};

enum class SpiResult
{
	Success = SPI_SUCCESS,
	NotImplement = SPI_NOT_IMPLEMENT,
	UserCancel = SPI_USER_CANCEL,
	UnknownFormat = SPI_UNKNOWN_FORMAT,
	DataBroken = SPI_DATA_BROKEN,
	NoMemory = SPI_NO_MEMORY,
	MemoryError = SPI_MEMORY_ERR,
	FileReadError = SPI_FILE_READ_ERR,
	Resered = SPI_RESERVED_ERR,
	InternalError = SPI_INTERNAL_ERR,
};

APTTYPE GetCurrentThreadApartmentType()
{
	APTTYPE           apttype;
	IComThreadingInfo *pComThreadingInfo;

	HRESULT hr = ::CoGetObjectContext(IID_PPV_ARGS(&pComThreadingInfo));

	hr = pComThreadingInfo->GetCurrentApartmentType(&apttype);

	hr = pComThreadingInfo->Release();
	return apttype;
}

inline HRESULT CreateMemoryStream(_Outptr_ IRandomAccessStream** stream)
{
	HRESULT hr;
	hr = Windows::Foundation::ActivateInstance(
		HStringReference(RuntimeClass_Windows_Storage_Streams_InMemoryRandomAccessStream).Get(),
		stream);
	return hr;
}

inline HRESULT CreateMemoryStream(_Outptr_ IStream** stream)
{
	HRESULT hr;
	ComPtr<IRandomAccessStream> abiStream;
	hr = Windows::Foundation::ActivateInstance(
		HStringReference(RuntimeClass_Windows_Storage_Streams_InMemoryRandomAccessStream).Get(),
		&abiStream);

	if (SUCCEEDED(hr))
	{
		hr = ::CreateStreamOverRandomAccessStream(abiStream.Get(), IID_PPV_ARGS(stream));
	}
	return hr;
}

HRESULT AsRandomAccessStream(IStream* src, _Outptr_ IRandomAccessStream **stream, BSOS_OPTIONS option = BSOS_DEFAULT)
{
	return CreateRandomAccessStreamOverStream(src, option, IID_PPV_ARGS(stream));
}

HRESULT AsStream(IRandomAccessStream *src, _Outptr_ IStream** stream)
{
	return CreateStreamOverRandomAccessStream(src, IID_PPV_ARGS(stream));
}

class Waiter
{
	Event m_event;

public:
	Waiter()
		: m_event(::CreateEventEx(NULL, NULL, CREATE_EVENT_MANUAL_RESET, EVENT_ALL_ACCESS))
	{
	}

	void Set()
	{
		SetEvent(m_event.Get());
	}
	HANDLE GetHandle()
	{
		return m_event.Get();
	}
	void Reset()
	{
		ResetEvent(m_event.Get());
	}
	operator Event&()
	{
		return m_event;
	}

	bool Wait(DWORD timeout = INFINITE)
	{
		return ::WaitForSingleObjectEx(m_event.Get(), timeout, FALSE) == WAIT_OBJECT_0;
	}
	DWORD get_Status()
	{
		DWORD dw = ::WaitForSingleObjectEx(m_event.Get(), 0, FALSE);
		return dw;
	}
	__declspec(property(get = get_Status)) DWORD Status;
};


HRESULT GetExtensionFromEncoder(REFGUID encoderId, WCHAR *fileExtensions, UINT cchfileExtensions)
{
	HRESULT hr;
	ComPtr<IWICBitmapEncoder> encoder;
	hr = ::CoCreateInstance(encoderId, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&encoder));

	if (FAILED(hr)) return hr;
	ComPtr<IWICBitmapEncoderInfo> encoderInfo;
	hr = encoder->GetEncoderInfo(&encoderInfo);
	if (FAILED(hr)) return hr;

	UINT cchActual;
	hr = encoderInfo->GetFileExtensions(cchfileExtensions, fileExtensions, &cchActual);
	if (FAILED(hr)) return hr;

	auto sep = wcschr(fileExtensions, L';');
	if (sep) *sep = L'\0';

	return S_OK;
}


int __stdcall GetPluginInfo(int infono, LPSTR buf, int buflen)
{
	if (buf == NULL || buflen <= 0)
	{
		return 0;
	}

	if (infono == 0)
	{
		return static_cast<int>(::strlen(::lstrcpynA(buf, "00AM", buflen)));
	}
	else if (infono == 1)
	{
		return static_cast<int>(::strlen(::lstrcpynA(buf, "axwrpdf build at "  __TIMESTAMP__, buflen)));
	}
	else if (infono == 2)
	{
		return static_cast<int>(::strlen(::lstrcpynA(buf, ".pdf", buflen)));
	}
	else if (infono == 3)
	{
		return static_cast<int>(::strlen(::lstrcpynA(buf, "PDF", buflen)));
	}

	return 0;
}

int __stdcall IsSupported(LPSTR filename, void* dw)
{
	if (memcmp(dw, "%PDF-", 5) == 0)
		return TRUE;

	return FALSE;
}

inline SpiResult HresultToSpiResult(HRESULT hr)
{
	switch (hr)
	{
	case S_OK:
	case S_FALSE:
		return SpiResult::Success;
	case E_NOTIMPL:
		return SpiResult::NotImplement;
	case E_OUTOFMEMORY:
		return SpiResult::NoMemory;
	case E_ABORT:
		return SpiResult::UserCancel;
	case E_ACCESSDENIED:
		return SpiResult::FileReadError;
	default:
		return SpiResult::InternalError;
	}
}

void GetPdfPage(IPdfPage *page, IStream *outStream, HANDLE hCompleteEvent)
{
	UINT i;
	CHECK_HR(page->get_Index(&i));
	auto callback = Callback<IAsyncActionCompletedHandler>([i, &hCompleteEvent](IAsyncAction* async, AsyncStatus status)
	{
		OutputDebugString((L"Start " + std::to_wstring(i) + L"\n").c_str());
		if (status != AsyncStatus::Completed)
		{
			return S_FALSE;
		}
		OutputDebugString((L"End " + std::to_wstring(i) + L"\n").c_str());
		::SetEvent(hCompleteEvent);
		return S_OK;
	});

	ComPtr<IRandomAccessStream> stream;
	CHECK_HR(AsRandomAccessStream(outStream, &stream));

	ComPtr<IAsyncAction> async;
	CHECK_HR(page->RenderToStreamAsync(stream.Get(), &async));
	CHECK_HR(async->put_Completed(callback.Get()));
}

inline void PdfPageToFileInfo(IPdfPage *page, UINT index, fileInfo &fileInfo, LPCSTR ext, SIZE_T filesize)
{
	memcpy(fileInfo.method, "PDF", 3);
	fileInfo.position = index;
	fileInfo.compsize = filesize;
	fileInfo.filesize = filesize;
	fileInfo.timestamp = static_cast<susie_time_t>(std::time(nullptr));
	sprintf_s(fileInfo.filename, "%08u%s", index, ext);
	fileInfo.crc = 0;
}
template<class TCallback>
HRESULT GetPdfDocument(LPCSTR buf, long len, bool isMemory, TCallback action)
{
	HRESULT hr;
	try
	{
		ComPtr<IPdfDocument> doc;
		ComPtr<IAsyncOperation<PdfDocument*>> async;
		Waiter w;
		{
			ComPtr<IRandomAccessStream> inStream;
			if (isMemory)
			{
				return E_NOTIMPL;
			}
			else
			{
				CHECK_HR(::CreateRandomAccessStreamOnFile(
					a2wstring(buf).c_str(), FileAccessMode_Read, IID_PPV_ARGS(&inStream)));
			}

			ComPtr<IPdfDocumentStatics> pdfDocumentsStatics;
			CHECK_HR(Windows::Foundation::GetActivationFactory(
				HStringReference(RuntimeClass_Windows_Data_Pdf_PdfDocument).Get(), &pdfDocumentsStatics));

			//return S_OK;
			CHECK_HR(pdfDocumentsStatics->LoadFromStreamAsync(inStream.Get(), &async));
		}
		auto callback = Callback<IAsyncOperationCompletedHandler<PdfDocument*>>(
			[&](_In_ IAsyncOperation<PdfDocument*>* pdfAsync, AsyncStatus status)
		{
			if (status == AsyncStatus::Started)
				return S_FALSE;

			std::shared_ptr<void> x(((Event&)w).Get(), ::SetEvent);

			if (status != AsyncStatus::Completed)
				return S_FALSE;

			try
			{
				CHECK_HR(pdfAsync->GetResults(&doc));

				bool useFile = true;
				ComPtr<IPdfPageRenderOptions> renderOptions;
				WCHAR ext[MAX_PATH];
				{
					CHECK_HR(Windows::Foundation::ActivateInstance(
						HStringReference(RuntimeClass_Windows_Data_Pdf_PdfPageRenderOptions).Get(), &renderOptions));
					GUID encoderId;
					CHECK_HR(renderOptions->get_BitmapEncoderId(&encoderId));
					CHECK_HR(GetExtensionFromEncoder(encoderId, ext, _countof(ext)));
				}

				hr = action(doc.Get(), ext);
				return hr;
			}
			catch (const _com_error &e)
			{
				return hr;
			}
		});
		CHECK_HR(async->put_Completed(callback.Get()));
		w.Wait(60 * 1000);

		//if (_p) _p->Release(), _p = nullptr;
		return hr;
	}
	catch (const _com_error &e)
	{
		return e.Error();
	}
}

HRESULT GetArchiveInfoInternal(LPCSTR buf, long len, unsigned int flag, HLOCAL * lphInf)
{
	if (lphInf == nullptr)
	{
		return E_INVALIDARG;
	}
	HRESULT hr = GetPdfDocument(buf, len, flag & 0b000111, [&](IPdfDocument *doc, const WCHAR(&ext)[MAX_PATH])
	{
		try
		{
			UINT count;
			CHECK_HR(doc->get_PageCount(&count));

			std::shared_ptr<fileInfo> fileList(new fileInfo[count + 1](), std::default_delete<fileInfo[]>());

			for (UINT i = 0; i < count; i++)
			{
				ComPtr<IPdfPage> page;
				CHECK_HR(doc->GetPage(i, &page));

				SIZE_T fileSize = 1;
				if (getAccurateSize)
				{
					Waiter w;
					ComPtr<IStream> memoryStream;
					CHECK_HR(::CreateStreamOnHGlobal(NULL, TRUE, &memoryStream));
					GetPdfPage(page.Get(), memoryStream.Get(), w.GetHandle());
					w.Wait();

					HGLOBAL hGlobal = nullptr;
					CHECK_HR(::GetHGlobalFromStream(memoryStream.Get(), &hGlobal));
					fileSize = ::GlobalSize(hGlobal);
				}

				PdfPageToFileInfo(page.Get(), i, fileList.get()[i], w2string(ext).c_str(), fileSize);
			}
			auto mem = ::LocalAlloc(LMEM_ZEROINIT, sizeof(fileInfo) * count + 1);
			memcpy(mem, fileList.get(), sizeof(fileInfo) * count + 1);

			*lphInf = mem;

			return S_OK;
		}
		catch (const _com_error &e)
		{
			return e.Error();
		}
		catch (const std::bad_alloc&)
		{
			return E_OUTOFMEMORY;
		}
	});

	//if (_p) _p->Release(), _p = nullptr;
	return hr;
}

int __stdcall GetArchiveInfo(LPCSTR buf, long len, unsigned int flag, HLOCAL * lphInf)
{
	HRESULT hr;
	try
	{
		std::thread worker([&] {
			RoInitializeWrapper init(RO_INIT_MULTITHREADED);
			hr = init;
			if (FAILED(hr)) return;

			hr = GetArchiveInfoInternal(buf, len, flag, lphInf);
			CoFreeUnusedLibrariesEx(0, 0);
		});
		worker.join();
	}
	catch (const _com_error &err)
	{
		hr = err.Error();
	}
	return (int)HresultToSpiResult(hr);
}
HRESULT GetFileInfoInternal(LPCSTR buf, long len, LPSTR filename, unsigned int flag, fileInfo * lpInfo)
{
	HRESULT hr = GetPdfDocument(buf, len, flag & 0b000111, [&](IPdfDocument *doc, const WCHAR(&ext)[MAX_PATH])
	{
		try
		{
			UINT count;
			CHECK_HR(doc->get_PageCount(&count));
			UINT i = len;
			if (len >= count)
			{
				return E_INVALIDARG;
			}

			ComPtr<ABI::Windows::Data::Pdf::IPdfPage> page;
			CHECK_HR(doc->GetPage(i, &page));

			if (i == count - 1)
			{
				REFCOUNT_DEBUG(page.Get());
				//page.Get()->AddRef();// TODO なぜか参照カウントを増やさないと終了時にアクセス違反が発生する。
				_p = page.Get();
				//workaroundRef = page.Get();
				REFCOUNT_DEBUG(page.Get());
			}

			SIZE_T fileSize = 1;
			if (getAccurateSize)
			{
				Waiter w;
				ComPtr<IStream> memoryStream;
				CHECK_HR(::CreateStreamOnHGlobal(NULL, TRUE, &memoryStream));
				GetPdfPage(page.Get(), memoryStream.Get(), w.GetHandle());
				w.Wait();

				HGLOBAL hGlobal = nullptr;
				CHECK_HR(::GetHGlobalFromStream(memoryStream.Get(), &hGlobal));
				fileSize = ::GlobalSize(hGlobal);
			}

			PdfPageToFileInfo(page.Get(), i, *lpInfo, w2string(ext).c_str(), fileSize);

			return S_OK;
		}
		catch (const _com_error &e)
		{
			hr = e.Error();
			return e.Error();
		}
		catch (const std::bad_alloc&)
		{
			return E_OUTOFMEMORY;
		}
	});

	//if (_p) _p->Release(), _p = nullptr;
	return hr;
}

int __stdcall GetFileInfo(LPCSTR buf, long len, LPSTR filename, unsigned int flag, fileInfo * lpInfo)
{
	HRESULT hr;
	try
	{
		std::thread worker([&] {
			RoInitializeWrapper init(RO_INIT_MULTITHREADED);
			hr = init;
			if (FAILED(hr)) return;

			hr = GetFileInfoInternal(buf, len, filename, flag, lpInfo);
			CoFreeUnusedLibrariesEx(0, 0);
		});
		worker.join();
	}
	catch (const std::exception&)
	{
		return (int)HresultToSpiResult(E_FAIL);
	}
	return (int)HresultToSpiResult(hr);
}

HRESULT GetFileInternal(LPCSTR buf, long len, LPSTR dest, unsigned int flag, FARPROC progressCallback, intptr_t lData)
{
	const bool useFile = (flag & 0b0000011100000000) == 0;
	HRESULT hr = GetPdfDocument(buf, len, flag & 0b000111, [&](IPdfDocument *doc, const WCHAR(&ext)[MAX_PATH])
	{
		try
		{
			UINT count;
			CHECK_HR(doc->get_PageCount(&count));
			UINT i = len;
			if (len >= count)
			{
				return E_INVALIDARG;
			}

			ComPtr<IStream> outStream;
			if (useFile)
			{
				WCHAR path[MAX_PATH];
				swprintf_s(path, L"%S\\%08u%s", buf, i, ext);
				CHECK_HR(::SHCreateStreamOnFileEx(path, STGM_WRITE | STGM_CREATE, FILE_ATTRIBUTE_NORMAL, TRUE, nullptr, &outStream));
			}
			else
			{
				CHECK_HR(::CreateStreamOnHGlobal(NULL, FALSE, &outStream));
			}
			Waiter w;
			ComPtr<ABI::Windows::Data::Pdf::IPdfPage> page;
			CHECK_HR(doc->GetPage(i, &page));
			GetPdfPage(page.Get(), outStream.Get(), w.GetHandle());

			if (i == count - 1)
			{
				REFCOUNT_DEBUG(page.Get());
				page.Get()->AddRef();// TODO なぜか参照カウントを増やさないと終了時にアクセス違反が発生する。
				_p = page.Get();
				//workaroundRef = page.Get();
				REFCOUNT_DEBUG(page.Get());
			}

			w.Wait();
			if (useFile)
			{
			}
			else
			{
				HGLOBAL hGlobal = nullptr;
				CHECK_HR(::GetHGlobalFromStream(outStream.Get(), &hGlobal));
				*reinterpret_cast<HLOCAL*>(dest) = hGlobal;
			}
			return S_OK;
		}
		catch (const _com_error &e)
		{
			hr = e.Error();
			return e.Error();
		}
		catch (const std::bad_alloc&)
		{
			return E_OUTOFMEMORY;
		}
	});

	//if (_p) _p->Release(), _p = nullptr;
	return hr;
}

int __stdcall GetFile(LPCSTR buf, long len, LPSTR dest, unsigned int flag, FARPROC progressCallback, intptr_t lData)
{
	HRESULT hr;
	try
	{
		std::thread worker([&] {
			RoInitializeWrapper init(RO_INIT_MULTITHREADED);
			hr = init;
			if (FAILED(hr)) return;

			hr = GetFileInternal(buf, len, dest, flag, progressCallback, lData);
			CoFreeUnusedLibrariesEx(0, 0);
		});
		worker.join();
	}
	catch (const std::exception&)
	{
		return (int)HresultToSpiResult(E_FAIL);
	}
	return (int)HresultToSpiResult(hr);
}

HRESULT PdfRender(LPCWSTR filePath)
{
	HRESULT hr;
	try
	{
		ComPtr<IPdfDocument> doc;
		ComPtr<IAsyncOperation<PdfDocument*>> async;
		Waiter w;
		{
			ComPtr<IRandomAccessStream> inStream;
			CHECK_HR(CreateRandomAccessStreamOnFile(
				filePath, FileAccessMode_Read, IID_PPV_ARGS(&inStream)));

			ComPtr<IPdfDocumentStatics> pdfDocumentsStatics;
			CHECK_HR(Windows::Foundation::GetActivationFactory(
				HStringReference(RuntimeClass_Windows_Data_Pdf_PdfDocument).Get(), &pdfDocumentsStatics));

			CHECK_HR(pdfDocumentsStatics->LoadFromStreamAsync(inStream.Get(), &async));
		}
		auto callback = Callback<IAsyncOperationCompletedHandler<PdfDocument*>>(
			[&](_In_ IAsyncOperation<PdfDocument*>* pdfAsync, AsyncStatus status)
		{
			if (status == AsyncStatus::Started)
				return S_FALSE;

			std::shared_ptr<void> x(((Event&)w).Get(), ::SetEvent);

			if (status != AsyncStatus::Completed)
				return S_FALSE;

			try
			{
				hr = pdfAsync->GetResults(&doc);

				bool useFile = true;
				UINT count;
				CHECK_HR(doc->get_PageCount(&count));

				struct PageInfo
				{
					UINT pageIndex;
					HGLOBAL hGlobal;
					ComPtr<IAsyncAction> async;
				};

				ComPtr<IPdfPageRenderOptions> renderOptions;
				WCHAR ext[MAX_PATH];
				{
					CHECK_HR(Windows::Foundation::ActivateInstance(
						HStringReference(RuntimeClass_Windows_Data_Pdf_PdfPageRenderOptions).Get(), &renderOptions));
					GUID encoderId;
					CHECK_HR(renderOptions->get_BitmapEncoderId(&encoderId));
					CHECK_HR(GetExtensionFromEncoder(encoderId, ext, _countof(ext)));
				}

				std::vector<PageInfo> pages(count);
				std::vector<HANDLE> waitevents(count);

				for (UINT i = count - 1; i < count; i++)
				{
					ComPtr<IStream> outStream;
					HGLOBAL hGlobal = nullptr;
					if (useFile)
					{
						CHECK_HR(::SHCreateStreamOnFileEx((L"page" + std::to_wstring(i) + L".png").c_str(), STGM_WRITE | STGM_CREATE, FILE_ATTRIBUTE_NORMAL, TRUE, nullptr, &outStream));
					}
					else
					{
						CHECK_HR(::CreateStreamOnHGlobal(NULL, FALSE, &outStream));
						CHECK_HR(::GetHGlobalFromStream(outStream.Get(), &hGlobal));
					}

					auto callback = Callback<IAsyncActionCompletedHandler>([&waitevents, i](IAsyncAction* async, AsyncStatus status)
					{
						std::shared_ptr<void> x(waitevents[i], ::SetEvent);
						OutputDebugString((L"Start " + std::to_wstring(i) + L"\n").c_str());
						if (status != AsyncStatus::Completed)
						{
							return S_FALSE;
						}
						OutputDebugString((L"End " + std::to_wstring(i) + L"\n").c_str());
						return S_OK;
					});
					waitevents[i] = ::CreateEventEx(nullptr, nullptr, 0, EVENT_ALL_ACCESS);

					ComPtr<IRandomAccessStream> stream;
					CHECK_HR(AsRandomAccessStream(outStream.Get(), &stream));

					ComPtr<ABI::Windows::Data::Pdf::IPdfPage> page;
					CHECK_HR(doc->GetPage(i, &page));
					ComPtr<IAsyncAction> async;
					CHECK_HR(page->RenderToStreamAsync(stream.Get(), &async));
					if (i == count - 1)
					{
						REFCOUNT_DEBUG(page.Get());
						page.Get()->AddRef();// TODO なぜか参照カウントを増やさないと終了時にアクセス違反が発生する。
						_p = page.Get();
						//workaroundRef = page.Get();
						REFCOUNT_DEBUG(page.Get());
					}
					CHECK_HR(async->put_Completed(callback.Get()));

					//pages.emplace_back(PageInfo{ i, hGlobal, std::move(async), });
				}

				::WaitForMultipleObjects(waitevents.size(), &waitevents.front(), TRUE, INFINITE);
				std::for_each(waitevents.begin(), waitevents.end(), [](HANDLE h) {::CloseHandle(h); });
				return hr;

				for (auto i : pages)
				{
					CHECK_HR(i.async->GetResults());
					if (useFile)
					{
					}
					else
					{
						assert(GlobalSize(i.hGlobal));
						auto p = (const BYTE*)::GlobalLock(i.hGlobal);
						printf("%02X %02X \n", p[0], p[1]);
						GlobalFree(i.hGlobal);
					}
				}
				return hr;
			}
			catch (const _com_error &e)
			{
				return hr;
			}
		});
		CHECK_HR(async->put_Completed(callback.Get()));
		w.Wait(60 * 1000);

		//if (_p) _p->Release(), _p = nullptr;
		return hr;
	}
	catch (const _com_error &e)
	{
		return e.Error();
	}
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int cmdShow)
{
	//ATLENSURE_SUCCEEDED(RoInitialize(RO_INIT_SINGLETHREADED));
	CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

	HRESULT hr;
	std::thread worker([&] {
		RoInitializeWrapper init(RO_INIT_MULTITHREADED);
		hr = init;
		if (FAILED(hr)) return;

		hr = PdfRender(L"sample.pdf");
		//CoFreeUnusedLibrariesEx(0, 0);
		Sleep(10000);
		//CoFreeUnusedLibrariesEx(0, 0);

	});
	worker.join();
	return 0;

}

