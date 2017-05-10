#ifndef _K_RICOH_MTP_H_
#define _K_RICOH_MTP_H_

#ifdef KRICOHMTPDLL_EXPORTS
#define K_RICOH_API __declspec(dllexport)
#else
#define K_RICOH_API __declspec(dllimport)
#endif

// MTP Header
//#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <iostream>
#include <string>
#include <sstream>
#include <PortableDevice.h>
#include <PortableDeviceApi.h>
#include <wpdmtpextensions.h>
#include <commdlg.h>
#include <atlstr.h>
#include <strsafe.h>
#include <wrl/client.h>
#include <list>

#define SELECTION_BUFFER_SIZE 81
#define RICOH_NAME "RICOH THETA S"
#define CLIENT_NAME         L"K_RICOH"
#define CLIENT_MAJOR_VER    1
#define CLIENT_MINOR_VER    0
#define CLIENT_REVISION     0

#define NUM_OBJECTS_TO_REQUEST  10

// MTP Library
#pragma comment(lib, "ws2_32.lib")
#pragma comment(lib, "ShlWapi.lib")
#pragma comment(lib, "PortableDeviceGuids.lib")

enum KRicohMTPError{
	COINITIALIZE_FAIL = 0,
	THERE_IS_NO_RICOH = 1, 
	CANNOT_OPEN_SESSION = 10,
	CANNOT_CLOSE_SESSION = 11,
	CANNOT_TAKE_PICTURE = 12,
	NO_RICOH_ERROR = 100
};

class K_RICOH_API KRicohMTP
{
public:
	KRicohMTP();
	virtual ~KRicohMTP();

private:
	Microsoft::WRL::ComPtr<IPortableDevice> device;
	enum KRicohMTPError last_error;

	// Private Methods
	bool IsRicoh(_In_ IPortableDeviceManager* deviceManager,
				_In_ PCWSTR pnpDeviceID);
	DWORD FindRicoh(__out int& ricoh_index);
	void GetClientInformation(_Outptr_result_maybenull_ IPortableDeviceValues** clientInformation);
	void GetRicohDevice(_Outptr_result_maybenull_ IPortableDevice** device);
	void RecursiveEnumerate(__in PCWSTR pszObjectID, __in IPortableDeviceContent* pContent, __out std::list<std::wstring>& deviceIDs);
	bool GetLastImageObjName(__in IPortableDevice* device, __out std::wstring& obj_name);
	HRESULT GetStringValue(__in IPortableDeviceProperties* pProperties, __in PCWSTR pszObjectID, __in REFPROPERTYKEY key, __out CAtlStringW& strStringValue);
	HRESULT StreamCopy(__out std::list<BYTE>& out_image, __in IStream* pSourceStream, __in DWORD cbTransferSize, __in DWORD* pcbWritten);
	void GetImage(__in IPortableDevice* device, __out std::list<BYTE>& out_image,
				__in const WCHAR* obj_name);
	void DeleteImage(__in IPortableDevice* device, __in const WCHAR* obj_name);

	HRESULT SendCommand(__in IPortableDevice* pDevice, __in WORD command, __out DWORD* result = NULL, 
						__in const ULONG* params = NULL, __in const int param_count = 0);
public:
	// if there is ricoh theta s, return true and set member, else return false
	bool InitRicohDevice();
	DWORD OpenSession(__in ULONG storage = 0x10001);
	DWORD CloseSession();
	DWORD TakePicture();
	bool GetOneImageAndDelete(__out std::list<BYTE>& out_image);
	int GetLastError();
};

#endif