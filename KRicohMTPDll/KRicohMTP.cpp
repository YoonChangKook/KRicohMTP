#include "KRicohMTP.h"

using namespace std;
using namespace Microsoft::WRL;

KRicohMTP::KRicohMTP()
	: last_error(KRicohMTPError::NO_RICOH_ERROR), device(nullptr)
{
	HRESULT hr = S_OK;

	// Enable the heap manager to terminate the process on heap error.
	HeapSetInformation(nullptr, HeapEnableTerminationOnCorruption, nullptr, 0);

	// Initialize COM for COINIT_MULTITHREADED
	hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (hr != S_OK)
		this->last_error = KRicohMTPError::COINITIALIZE_FAIL;
}

KRicohMTP::~KRicohMTP()
{
	if (this->device != nullptr)
	{
		device->Close();
		//device->Release();
	}
}

bool KRicohMTP::InitRicohDevice()
{
	GetRicohDevice(&this->device);

	if (this->device == nullptr)
		return false;
	else
		return true;
}

DWORD KRicohMTP::OpenSession(__in ULONG storage)
{
	DWORD result = 0x2002;
	if (this->device == nullptr)
	{
		this->last_error = KRicohMTPError::THERE_IS_NO_RICOH;
		return result;
	}

	ULONG params[1] = { storage };
	if (SendCommand(this->device.Get(), 0x1002, &result, params, 1) != S_OK)
	{
		// ERROR
		this->last_error = KRicohMTPError::CANNOT_OPEN_SESSION;
	}

	Sleep(500);

	return result;
}

DWORD KRicohMTP::CloseSession()
{
	DWORD result = 0x2002;
	if (this->device == nullptr)
	{
		this->last_error = KRicohMTPError::THERE_IS_NO_RICOH;
		return result;
	}

	if (SendCommand(this->device.Get(), 0x1003, &result) != S_OK)
	{
		// ERROR
		this->last_error = KRicohMTPError::CANNOT_CLOSE_SESSION;
	}

	Sleep(500);

	return result;
}

DWORD KRicohMTP::TakePicture()
{
	DWORD result = 0x2002;
	if (this->device == nullptr)
	{
		this->last_error = KRicohMTPError::THERE_IS_NO_RICOH;
		return result;
	}

	if (SendCommand(this->device.Get(), 0x100E, &result) != S_OK)
	{
		// ERROR
		this->last_error = KRicohMTPError::CANNOT_TAKE_PICTURE;
	}
	
	Sleep(10000);

	return result;
}

bool KRicohMTP::GetOneImageAndDelete(__out std::list<BYTE>& out_image)
{
	std::wstring last_picture_id;

	if (this->device == nullptr)
	{
		this->last_error = KRicohMTPError::THERE_IS_NO_RICOH;
		return false;
	}

	// Get Image
	if (!GetLastImageObjName(this->device.Get(), last_picture_id))
	{
		return false;
	}

	Sleep(500);

	// Image Copy
	GetImage(this->device.Get(), out_image, last_picture_id.c_str());

	Sleep(500);

	// Delete Pictures
	DeleteImage(this->device.Get(), last_picture_id.c_str());

	return true;
}

int KRicohMTP::GetLastError()
{
	return this->last_error;
}

// Private Methods
bool KRicohMTP::IsRicoh(_In_ IPortableDeviceManager* deviceManager,
						_In_ PCWSTR pnpDeviceID)
{
	DWORD descriptionLength = 0;
	bool is_ricoh = false;

	// 1) Pass nullptr as the PWSTR return string parameter to get the total number
	// of characters to allocate for the string value.
	HRESULT hr = deviceManager->GetDeviceDescription(pnpDeviceID, nullptr, &descriptionLength);
	if (FAILED(hr))
	{
		wprintf(L"! Failed to get number of characters for device description, hr = 0x%lx\n", hr);
		return false;
	}
	else if (descriptionLength > 0)
	{
		// 2) Allocate the number of characters needed and retrieve the string value.
		PWSTR description = new (std::nothrow) WCHAR[descriptionLength];
		if (description != nullptr)
		{
			ZeroMemory(description, descriptionLength * sizeof(WCHAR));
			hr = deviceManager->GetDeviceDescription(pnpDeviceID, description, &descriptionLength);
			if (SUCCEEDED(hr))
			{
				char* temp_description = new char[descriptionLength];
				WideCharToMultiByte(CP_ACP,
					0,
					description,
					-1,
					temp_description,
					descriptionLength,
					NULL,
					NULL);

				if (strcmp(temp_description, RICOH_NAME) == 0)
					is_ricoh = true;

				delete[] temp_description;
				//wprintf(L"Description:   %ws\n", description);
			}
			else
			{
				wprintf(L"! Failed to get device description, hr = 0x%lx\n", hr);
			}

			// Delete the allocated description string
			delete[] description;
			description = nullptr;
		}
		else
		{
			wprintf(L"! Failed to allocate memory for the device description string\n");
		}

		return is_ricoh;
	}
	else
	{
		wprintf(L"The device did not provide a description.\n");
		return false;
	}
}

DWORD KRicohMTP::FindRicoh(__out int& ricoh_index)
{
	DWORD                           pnpDeviceIDCount = 0;
	ComPtr<IPortableDeviceManager>  deviceManager;
	ricoh_index = -1;

	// CoCreate the IPortableDeviceManager interface to enumerate
	// portable devices and to get information about them.
	//<SnippetDeviceEnum1>
	HRESULT hr = CoCreateInstance(CLSID_PortableDeviceManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&deviceManager));
	if (FAILED(hr))
	{
		wprintf(L"! Failed to CoCreateInstance CLSID_PortableDeviceManager, hr = 0x%lx\n", hr);
	}
	//</SnippetDeviceEnum1>

	// 1) Pass nullptr as the PWSTR array pointer to get the total number
	// of devices found on the system.
	//<SnippetDeviceEnum2>
	if (SUCCEEDED(hr))
	{
		hr = deviceManager->GetDevices(nullptr, &pnpDeviceIDCount);
		if (FAILED(hr))
		{
			wprintf(L"! Failed to get number of devices on the system, hr = 0x%lx\n", hr);
		}
	}

	// Report the number of devices found.  NOTE: we will report 0, if an error
	// occured.

	wprintf(L"\n%u Windows Portable Device(s) found on the system\n\n", pnpDeviceIDCount);
	//</SnippetDeviceEnum2>
	// 2) Allocate an array to hold the PnPDeviceID strings returned from
	// the IPortableDeviceManager::GetDevices method
	//<SnippetDeviceEnum3>
	if (SUCCEEDED(hr) && (pnpDeviceIDCount > 0))
	{
		PWSTR* pnpDeviceIDs = new (std::nothrow) PWSTR[pnpDeviceIDCount];
		if (pnpDeviceIDs != nullptr)
		{
			ZeroMemory(pnpDeviceIDs, pnpDeviceIDCount * sizeof(PWSTR));
			DWORD retrievedDeviceIDCount = pnpDeviceIDCount;
			hr = deviceManager->GetDevices(pnpDeviceIDs, &retrievedDeviceIDCount);
			if (SUCCEEDED(hr))
			{
				_Analysis_assume_(retrievedDeviceIDCount <= pnpDeviceIDCount);
				// For each device found, display the devices friendly name,
				// manufacturer, and description strings.
				for (DWORD index = 0; index < retrievedDeviceIDCount; index++)
				{
					if (IsRicoh(deviceManager.Get(), pnpDeviceIDs[index]) == true)
						ricoh_index = index;
				}
			}
			else
			{
				wprintf(L"! Failed to get the device list from the system, hr = 0x%lx\n", hr);
			}
			//</SnippetDeviceEnum3>

			// Free all returned PnPDeviceID strings by using CoTaskMemFree.
			// NOTE: CoTaskMemFree can handle nullptr pointers, so no nullptr
			//       check is needed.
			for (DWORD index = 0; index < pnpDeviceIDCount; index++)
			{
				CoTaskMemFree(pnpDeviceIDs[index]);
				pnpDeviceIDs[index] = nullptr;
			}

			// Delete the array of PWSTR pointers
			delete[] pnpDeviceIDs;
			pnpDeviceIDs = nullptr;
		}
		else
		{
			wprintf(L"! Failed to allocate memory for PWSTR array\n");
		}
	}

	return pnpDeviceIDCount;
}

void KRicohMTP::GetClientInformation(_Outptr_result_maybenull_ IPortableDeviceValues** clientInformation)
{
	// Client information is optional.  The client can choose to identify itself, or
	// to remain unknown to the driver.  It is beneficial to identify yourself because
	// drivers may be able to optimize their behavior for known clients. (e.g. An
	// IHV may want their bundled driver to perform differently when connected to their
	// bundled software.)

	// CoCreate an IPortableDeviceValues interface to hold the client information.
	HRESULT hr = CoCreateInstance(CLSID_PortableDeviceValues,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(clientInformation));
	if (SUCCEEDED(hr))
	{
		// Attempt to set all bits of client information
		hr = (*clientInformation)->SetStringValue(WPD_CLIENT_NAME, CLIENT_NAME);
		if (FAILED(hr))
		{
			wprintf(L"! Failed to set WPD_CLIENT_NAME, hr = 0x%lx\n", hr);
		}

		hr = (*clientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_MAJOR_VERSION, CLIENT_MAJOR_VER);
		if (FAILED(hr))
		{
			wprintf(L"! Failed to set WPD_CLIENT_MAJOR_VERSION, hr = 0x%lx\n", hr);
		}

		hr = (*clientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_MINOR_VERSION, CLIENT_MINOR_VER);
		if (FAILED(hr))
		{
			wprintf(L"! Failed to set WPD_CLIENT_MINOR_VERSION, hr = 0x%lx\n", hr);
		}

		hr = (*clientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_REVISION, CLIENT_REVISION);
		if (FAILED(hr))
		{
			wprintf(L"! Failed to set WPD_CLIENT_REVISION, hr = 0x%lx\n", hr);
		}

		//  Some device drivers need to impersonate the caller in order to function correctly.  Since our application does not
		//  need to restrict its identity, specify SECURITY_IMPERSONATION so that we work with all devices.
		hr = (*clientInformation)->SetUnsignedIntegerValue(WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE, SECURITY_IMPERSONATION);
		if (FAILED(hr))
		{
			wprintf(L"! Failed to set WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE, hr = 0x%lx\n", hr);
		}
	}
	else
	{
		wprintf(L"! Failed to CoCreateInstance CLSID_PortableDeviceValues, hr = 0x%lx\n", hr);
	}
}

void KRicohMTP::GetRicohDevice(_Outptr_result_maybenull_ IPortableDevice** device)
{
	*device = nullptr;

	HRESULT							hr = S_OK;
	//UINT   	                        currentDeviceIndex = 0;
	DWORD  	                        pnpDeviceIDCount = 0;
	PWSTR* 	                        pnpDeviceIDs = nullptr;
	WCHAR  	                        selection[SELECTION_BUFFER_SIZE] = { 0 };
	int	   							ricoh_index = -1;

	ComPtr<IPortableDeviceManager>	deviceManager;
	ComPtr<IPortableDeviceValues> 	clientInformation;

	// Fill out information about your application, so the device knows
	// who they are speaking to.

	GetClientInformation(&clientInformation);

	// Enumerate and display all devices.
	pnpDeviceIDCount = FindRicoh(ricoh_index);

	if (ricoh_index < 0)
	{
		return;
	}

	// CoCreate the IPortableDeviceManager interface to enumerate
	// portable devices and to get information about them.
	hr = CoCreateInstance(CLSID_PortableDeviceManager,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&deviceManager));
	if (FAILED(hr))
	{
		wprintf(L"! Failed to CoCreateInstance CLSID_PortableDeviceManager, hr = 0x%lx\n", hr);
	}

	// Allocate an array to hold the PnPDeviceID strings returned from
	// the IPortableDeviceManager::GetDevices method
	if (SUCCEEDED(hr) && (pnpDeviceIDCount > 0))
	{
		pnpDeviceIDs = new (std::nothrow)PWSTR[pnpDeviceIDCount];
		if (pnpDeviceIDs != nullptr)
		{
			ZeroMemory(pnpDeviceIDs, pnpDeviceIDCount * sizeof(PWSTR));
			DWORD retrievedDeviceIDCount = pnpDeviceIDCount;
			hr = deviceManager->GetDevices(pnpDeviceIDs, &retrievedDeviceIDCount);
			if (SUCCEEDED(hr))
			{
				_Analysis_assume_(retrievedDeviceIDCount <= pnpDeviceIDCount);
				// CoCreate the IPortableDevice interface and call Open() with
				// the chosen PnPDeviceID string.
				hr = CoCreateInstance(CLSID_PortableDeviceFTM,
					nullptr,
					CLSCTX_INPROC_SERVER,
					IID_PPV_ARGS(device));
				if (SUCCEEDED(hr))
				{
					hr = (*device)->Open(pnpDeviceIDs[ricoh_index], clientInformation.Get());

					if (hr == E_ACCESSDENIED)
					{
						wprintf(L"Failed to Open the device for Read Write access, will open it for Read-only access instead\n");
						clientInformation->SetUnsignedIntegerValue(WPD_CLIENT_DESIRED_ACCESS, GENERIC_READ);
						hr = (*device)->Open(pnpDeviceIDs[ricoh_index], clientInformation.Get());
					}

					if (FAILED(hr))
					{
						wprintf(L"! Failed to Open the device, hr = 0x%lx\n", hr);
						// Release the IPortableDevice interface, because we cannot proceed
						// with an unopen device.
						(*device)->Release();
						*device = nullptr;
					}
				}
				else
				{
					wprintf(L"! Failed to CoCreateInstance CLSID_PortableDeviceFTM, hr = 0x%lx\n", hr);
				}
			}
			else
			{
				wprintf(L"! Failed to get the device list from the system, hr = 0x%lx\n", hr);
			}

			// Free all returned PnPDeviceID strings by using CoTaskMemFree.
			// NOTE: CoTaskMemFree can handle nullptr pointers, so no nullptr
			//       check is needed.
			for (DWORD index = 0; index < pnpDeviceIDCount; index++)
			{
				CoTaskMemFree(pnpDeviceIDs[index]);
				pnpDeviceIDs[index] = nullptr;
			}

			// Delete the array of PWSTR pointers
			delete[] pnpDeviceIDs;
			pnpDeviceIDs = nullptr;
		}
		else
		{
			wprintf(L"! Failed to allocate memory for PWSTR array\n");
		}
	}

	// If no devices were found on the system, just exit this function.
}

bool KRicohMTP::GetLastImageObjName(__in IPortableDevice* device, __out std::wstring& obj_name)
{
	HRESULT                         hr = S_OK;
	ComPtr<IPortableDeviceContent> pContent;
	std::list<std::wstring> contentIDs;
	std::list<std::string> pictureIDs;
	PCWSTR picture_pre = L"o64";

	if (device == NULL)
	{
		printf("! A NULL IPortableDevice interface pointer was received\n");
		return false;
	}

	// Get an IPortableDeviceContent interface from the IPortableDevice interface to
	// access the content-specific methods.
	hr = device->Content(&pContent);
	if (FAILED(hr))
	{
		printf("! Failed to get IPortableDeviceContent from IPortableDevice, hr = 0x%lx\n", hr);
		return false;
	}

	// Enumerate content starting from the "DEVICE" object.
	if (SUCCEEDED(hr))
	{
		printf("\n");
		RecursiveEnumerate(WPD_DEVICE_OBJECT_ID, pContent.Get(), contentIDs);
	}

	wstring pre_wstr(picture_pre);
	std::string pre_str(pre_wstr.begin(), pre_wstr.end());
	for (std::list<std::wstring>::iterator it = contentIDs.begin(); it != contentIDs.end(); it++)
	{
		std::string content_str(it->begin(), it->end());
		// if the content is an image
		if (content_str.find(pre_str) != std::string::npos)
			pictureIDs.push_back(content_str);
	}

	// Get Last Picture
	int index = -1;
	std::string last_picture_id;
	for (std::list<std::string>::iterator it = pictureIDs.begin(); it != pictureIDs.end(); it++)
	{
		char buffer[4];
		std::string id_str = "0x";
		int id;
		it->copy(buffer, 4, 3);
		id_str += buffer;
		id = std::stol(id_str, nullptr, 16);

		if (index < id && id != 0)
		{
			index = id;
			last_picture_id = *it;
		}
	}

	obj_name = std::wstring(last_picture_id.begin(), last_picture_id.end());

	if (index > 0)
		return true;
	else
		return false;
}

HRESULT KRicohMTP::GetStringValue(__in IPortableDeviceProperties* pProperties,
								__in PCWSTR                     pszObjectID,
								__in REFPROPERTYKEY             key,
								__out CAtlStringW&              strStringValue)
{
	CComPtr<IPortableDeviceValues>        pObjectProperties;
	CComPtr<IPortableDeviceKeyCollection> pPropertiesToRead;

	// 1) CoCreate an IPortableDeviceKeyCollection interface to hold the the property key
	// we wish to read.
	HRESULT hr = CoCreateInstance(CLSID_PortableDeviceKeyCollection,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&pPropertiesToRead));

	// 2) Populate the IPortableDeviceKeyCollection with the keys we wish to read.
	// NOTE: We are not handling any special error cases here so we can proceed with
	// adding as many of the target properties as we can.
	if (SUCCEEDED(hr))
	{
		if (pPropertiesToRead != NULL)
		{
			HRESULT hrTemp = S_OK;
			hrTemp = pPropertiesToRead->Add(key);
			if (FAILED(hrTemp))
			{
				printf("! Failed to add PROPERTYKEY to IPortableDeviceKeyCollection, hr= 0x%lx\n", hrTemp);
			}
		}
	}

	// 3) Call GetValues() passing the collection of specified PROPERTYKEYs.
	if (SUCCEEDED(hr))
	{
		hr = pProperties->GetValues(pszObjectID,         // The object whose properties we are reading
			pPropertiesToRead,   // The properties we want to read
			&pObjectProperties); // Driver supplied property values for the specified object

		// The error is handled by the caller, which also displays an error message to the user.
	}

	// 4) Extract the string value from the returned property collection
	if (SUCCEEDED(hr))
	{
		PWSTR pszStringValue = NULL;
		hr = pObjectProperties->GetStringValue(key, &pszStringValue);
		if (SUCCEEDED(hr))
		{
			// Assign the newly read string to the CAtlStringW out
			// parameter.
			strStringValue = pszStringValue;
		}
		else
		{
			printf("! Failed to find property in IPortableDeviceValues, hr = 0x%lx\n", hr);
		}

		CoTaskMemFree(pszStringValue);
		pszStringValue = NULL;
	}

	return hr;
}

HRESULT KRicohMTP::StreamCopy(__out std::list<BYTE>& out_image, __in IStream* pSourceStream, __in DWORD cbTransferSize, __in DWORD* pcbWritten)
{
	HRESULT hr = S_OK;

	// Allocate a temporary buffer (of Optimal transfer size) for the read results to
	// be written to.
	BYTE*   pObjectData = new (std::nothrow) BYTE[cbTransferSize];
	if (pObjectData != NULL)
	{
		DWORD cbTotalBytesRead = 0;
		DWORD cbTotalBytesWritten = 0;

		DWORD cbBytesRead = 0;
		DWORD cbBytesWritten = 0;

		// Read until the number of bytes returned from the source stream is 0, or
		// an error occured during transfer.
		do
		{
			// Read object data from the source stream
			hr = pSourceStream->Read(pObjectData, cbTransferSize, &cbBytesRead);
			
			if (FAILED(hr))
			{
				printf("! Failed to read %d bytes from the source stream, hr = 0x%lx\n", cbTransferSize, hr);
			}

			// Write object data to the destination stream
			if (SUCCEEDED(hr))
			{
				cbTotalBytesRead += cbBytesRead; // Calculating total bytes read from device for debugging purposes only

				cbBytesWritten = cbBytesRead;
				for (int i = 0; i < cbBytesRead; i++)
					out_image.push_back(pObjectData[i]);
				//hr = pDestStream->Write(pObjectData, cbBytesRead, &cbBytesWritten);
				//if (FAILED(hr))
				//{
				//	printf("! Failed to write %d bytes of object data to the destination stream, hr = 0x%lx\n", cbBytesRead, hr);
				//}
				//
				//if (SUCCEEDED(hr))
				//{
				//	cbTotalBytesWritten += cbBytesWritten; // Calculating total bytes written to the file for debugging purposes only
				//}
			}

			// Output Read/Write operation information only if we have received data and if no error has occured so far.
			if (SUCCEEDED(hr) && (cbBytesRead > 0))
			{
				printf("Read %d bytes from the source stream...Wrote %d bytes to the destination stream...\n", cbBytesRead, cbBytesWritten);
			}

		} while (SUCCEEDED(hr) && (cbBytesRead > 0));

		// If the caller supplied a pcbWritten parameter and we
		// and we are successful, set it to cbTotalBytesWritten
		// before exiting.
		if ((SUCCEEDED(hr)) && (pcbWritten != NULL))
		{
			*pcbWritten = cbTotalBytesWritten;
		}

		// Remember to delete the temporary transfer buffer
		delete[] pObjectData;
		pObjectData = NULL;
	}
	else
	{
		printf("! Failed to allocate %d bytes for the temporary transfer buffer.\n", cbTransferSize);
	}

	return hr;
}

void KRicohMTP::GetImage(__in IPortableDevice* device, __out std::list<BYTE>& out_image, __in const WCHAR* obj_name)
{
	HRESULT								hr = S_OK;
	ComPtr<IPortableDeviceContent>		pContent;
	ComPtr<IPortableDeviceResources>	pResources;
	ComPtr<IPortableDeviceProperties>	pProperties;
	ComPtr<IStream>						pObjectDataStream;
	ComPtr<IStream>						pFinalFileStream;
	DWORD								cbOptimalTransferSize = 0;
	CAtlStringW							strOriginalFileName;

	if (device == NULL)
	{
		printf("! A NULL IPortableDevice interface pointer was received\n");
		return;
	}

	//</SnippetTransferFrom1>
	// 1) get an IPortableDeviceContent interface from the IPortableDevice interface to
	// access the content-specific methods.
	//<SnippetTransferFrom2>
	if (SUCCEEDED(hr))
	{
		hr = device->Content(&pContent);
		if (FAILED(hr))
		{
			printf("! Failed to get IPortableDeviceContent from IPortableDevice, hr = 0x%lx\n", hr);
		}
	}
	//</SnippetTransferFrom2>
	// 2) Get an IPortableDeviceResources interface from the IPortableDeviceContent interface to
	// access the resource-specific methods.
	//<SnippetTransferFrom3>
	if (SUCCEEDED(hr))
	{
		hr = pContent->Transfer(&pResources);
		if (FAILED(hr))
		{
			printf("! Failed to get IPortableDeviceResources from IPortableDeviceContent, hr = 0x%lx\n", hr);
		}
	}
	//</SnippetTransferFrom3>
	// 3) Get the IStream (with READ access) and the optimal transfer buffer size
	// to begin the transfer.
	//<SnippetTransferFrom4>
	if (SUCCEEDED(hr))
	{
		hr = pResources->GetStream(obj_name,             // Identifier of the object we want to transfer
								WPD_RESOURCE_DEFAULT,    // We are transferring the default resource (which is the entire object's data)
								STGM_READ,               // Opening a stream in READ mode, because we are reading data from the device.
								&cbOptimalTransferSize,  // Driver supplied optimal transfer size
								&pObjectDataStream);
		if (FAILED(hr))
		{
			printf("! Failed to get IStream (representing object data on the device) from IPortableDeviceResources, hr = 0x%lx\n", hr);
		}
	}
	//</SnippetTransferFrom4>

	// 4) Read the WPD_OBJECT_ORIGINAL_FILE_NAME property so we can properly name the
	// transferred object.  Some content objects may not have this property, so a
	// fall-back case has been provided below. (i.e. Creating a file named <objectID>.data )
	//<SnippetTransferFrom5>
	if (SUCCEEDED(hr))
	{
		hr = pContent->Properties(&pProperties);
		if (SUCCEEDED(hr))
		{
			hr = GetStringValue(pProperties.Get(),
								obj_name,
								WPD_OBJECT_ORIGINAL_FILE_NAME,
								strOriginalFileName);
			if (FAILED(hr))
			{
				printf("! Failed to read WPD_OBJECT_ORIGINAL_FILE_NAME on object '%ws', hr = 0x%lx\n", obj_name, hr);
				strOriginalFileName.Format(L"%ws.data", obj_name);
				printf("* Creating a filename '%ws' as a default.\n", (PWSTR)strOriginalFileName.GetString());
				// Set the HRESULT to S_OK, so we can continue with our newly generated
				// temporary file name.
				hr = S_OK;
			}
		}
		else
		{
			printf("! Failed to get IPortableDeviceProperties from IPortableDeviceContent, hr = 0x%lx\n", hr);
		}
	}

	//</SnippetTransferFrom5>
	// 5) Read on the object's data stream and write to the final file's data stream using the
    // driver supplied optimal transfer buffer size.
	//<SnippetTransferFrom6>
	if (SUCCEEDED(hr))
	{
		DWORD cbTotalBytesWritten = 0;

		// Since we have IStream-compatible interfaces, call our helper function
		// that copies the contents of a source stream into a destination stream.
		hr = StreamCopy(out_image,      // Destination (The Final File to transfer to)
			pObjectDataStream.Get(),    // Source (The Object's data to transfer from)
			cbOptimalTransferSize,		// The driver specified optimal transfer buffer size
			&cbTotalBytesWritten);		// The total number of bytes transferred from device to the finished file
		if (FAILED(hr))
		{
			printf("! Failed to transfer object from device, hr = 0x%lx\n", hr);
		}
		else
		{
			printf("* Transferred object '%ws' to '%s'.\n", obj_name, "std::list<BYTE> out_image");
		}
	}
}

void KRicohMTP::DeleteImage(__in IPortableDevice* device, __in const WCHAR* obj_name)
{
	HRESULT                                       hr = S_OK;
	CComPtr<IPortableDeviceContent>               pContent;
	CComPtr<IPortableDevicePropVariantCollection> pObjectsToDelete;
	CComPtr<IPortableDevicePropVariantCollection> pObjectsFailedToDelete;

	if (device == NULL)
	{
		printf("! A NULL IPortableDevice interface pointer was received\n");
		return;
	}

	// 1) get an IPortableDeviceContent interface from the IPortableDevice interface to
	// access the content-specific methods.
	if (SUCCEEDED(hr))
	{
		hr = device->Content(&pContent);
		if (FAILED(hr))
		{
			printf("! Failed to get IPortableDeviceContent from IPortableDevice, hr = 0x%lx\n", hr);
		}
	}

	// 2) CoCreate an IPortableDevicePropVariantCollection interface to hold the the object identifiers
	// to delete.
	//
	// NOTE: This is a collection interface so more than 1 object can be deleted at a time.
	//       This sample only deletes a single object.
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(CLSID_PortableDevicePropVariantCollection,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_PPV_ARGS(&pObjectsToDelete));
		if (SUCCEEDED(hr))
		{
			if (pObjectsToDelete != NULL)
			{
				PROPVARIANT pv = { 0 };
				PropVariantInit(&pv);

				// Initialize a PROPVARIANT structure with the object identifier string
				// that the user selected above. Notice we are allocating memory for the
				// PWSTR value.  This memory will be freed when PropVariantClear() is
				// called below.
				pv.vt = VT_LPWSTR;
				pv.pwszVal = AtlAllocTaskWideString(obj_name);
				if (pv.pwszVal != NULL)
				{
					// Add the object identifier to the objects-to-delete list
					// (We are only deleting 1 in this example)
					hr = pObjectsToDelete->Add(&pv);
					if (SUCCEEDED(hr))
					{
						// Attempt to delete the object from the device
						hr = pContent->Delete(PORTABLE_DEVICE_DELETE_NO_RECURSION,  // Deleting with no recursion
							pObjectsToDelete,                     // Object(s) to delete
							NULL);                                // Object(s) that failed to delete (we are only deleting 1, so we can pass NULL here)
						if (SUCCEEDED(hr))
						{
							// An S_OK return lets the caller know that the deletion was successful
							if (hr == S_OK)
							{
								printf("The object '%ws' was deleted from the device.\n", obj_name);
							}

							// An S_FALSE return lets the caller know that the deletion failed.
							// The caller should check the returned IPortableDevicePropVariantCollection
							// for a list of object identifiers that failed to be deleted.
							else
							{
								printf("The object '%ws' failed to be deleted from the device.\n", obj_name);
							}
						}
						else
						{
							printf("! Failed to delete an object from the device, hr = 0x%lx\n", hr);
						}
					}
					else
					{
						printf("! Failed to delete an object from the device because we could no add the object identifier string to the IPortableDevicePropVariantCollection, hr = 0x%lx\n", hr);
					}
				}
				else
				{
					hr = E_OUTOFMEMORY;
					printf("! Failed to delete an object from the device because we could no allocate memory for the object identifier string, hr = 0x%lx\n", hr);
				}

				// Free any allocated values in the PROPVARIANT before exiting
				PropVariantClear(&pv);
			}
			else
			{
				printf("! Failed to delete an object from the device because we were returned a NULL IPortableDevicePropVariantCollection interface pointer, hr = 0x%lx\n", hr);
			}
		}
		else
		{
			printf("! Failed to CoCreateInstance CLSID_PortableDevicePropVariantCollection, hr = 0x%lx\n", hr);
		}
	}
}

void KRicohMTP::RecursiveEnumerate(__in PCWSTR pszObjectID, __in IPortableDeviceContent* pContent, __out std::list<std::wstring>& deviceIDs)
{
	ComPtr<IEnumPortableDeviceObjectIDs> pEnumObjectIDs;

	// Print the object identifier being used as the parent during enumeration.
	//printf("%ws\n", pszObjectID);
	deviceIDs.push_back(std::wstring(pszObjectID));

	// Get an IEnumPortableDeviceObjectIDs interface by calling EnumObjects with the
	// specified parent object identifier.
	HRESULT hr = pContent->EnumObjects(0,               // Flags are unused
		pszObjectID,     // Starting from the passed in object
		NULL,            // Filter is unused
		&pEnumObjectIDs);
	if (FAILED(hr))
	{
		printf("! Failed to get IEnumPortableDeviceObjectIDs from IPortableDeviceContent, hr = 0x%lx\n", hr);
	}

	// Loop calling Next() while S_OK is being returned.
	while (hr == S_OK)
	{
		DWORD  cFetched = 0;
		PWSTR  szObjectIDArray[NUM_OBJECTS_TO_REQUEST] = { 0 };
		hr = pEnumObjectIDs->Next(NUM_OBJECTS_TO_REQUEST,   // Number of objects to request on each NEXT call
			szObjectIDArray,          // Array of PWSTR array which will be populated on each NEXT call
			&cFetched);               // Number of objects written to the PWSTR array
		if (SUCCEEDED(hr))
		{
			// Traverse the results of the Next() operation and recursively enumerate
			// Remember to free all returned object identifiers using CoTaskMemFree()
			for (DWORD dwIndex = 0; dwIndex < cFetched; dwIndex++)
			{
				RecursiveEnumerate(szObjectIDArray[dwIndex], pContent, deviceIDs);

				// Free allocated PWSTRs after the recursive enumeration call has completed.
				CoTaskMemFree(szObjectIDArray[dwIndex]);
				szObjectIDArray[dwIndex] = NULL;
			}
		}
	}
}

HRESULT KRicohMTP::SendCommand(__in IPortableDevice* pDevice, __in WORD command, __out DWORD* result,
							__in const ULONG* params, __in const int param_count)
{
	HRESULT hr = S_OK;
	const WORD PTP_OPCODE_GETNUMOBJECT = command; // GetNumObject opcode is 0x1006
	const WORD PTP_RESPONSECODE_OK = 0x2001;     // 0x2001 indicates command success

	// Build basic WPD parameters for the command
	ComPtr<IPortableDeviceValues> spParameters;
	if (hr == S_OK)
	{
		hr = CoCreateInstance(CLSID_PortableDeviceValues,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_IPortableDeviceValues,
			(VOID**)&spParameters);
	}

	// Use the WPD_COMMAND_MTP_EXT_EXECUTE_COMMAND_WITHOUT_DATA_PHASE command
	// Similar commands exist for reading and writing data phases
	if (hr == S_OK)
	{
		hr = spParameters->SetGuidValue(WPD_PROPERTY_COMMON_COMMAND_CATEGORY,
			WPD_COMMAND_MTP_EXT_EXECUTE_COMMAND_WITHOUT_DATA_PHASE.fmtid);
	}

	if (hr == S_OK)
	{
		hr = spParameters->SetUnsignedIntegerValue(WPD_PROPERTY_COMMON_COMMAND_ID,
			WPD_COMMAND_MTP_EXT_EXECUTE_COMMAND_WITHOUT_DATA_PHASE.pid);
	}

	// Specify the actual MTP opcode that we want to execute here
	if (hr == S_OK)
	{
		hr = spParameters->SetUnsignedIntegerValue(WPD_PROPERTY_MTP_EXT_OPERATION_CODE,
			(ULONG)PTP_OPCODE_GETNUMOBJECT);
	}

	// GetNumObject requires three params - storage ID, object format, and parent object handle   
	// Parameters need to be first put into a PropVariantCollection
	ComPtr<IPortableDevicePropVariantCollection> spMtpParams;
	if (hr == S_OK)
	{
		hr = CoCreateInstance(CLSID_PortableDevicePropVariantCollection,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_IPortableDevicePropVariantCollection,
			(VOID**)&spMtpParams);
	}

	PROPVARIANT pvParam = { 0 };
	pvParam.vt = VT_UI4;

	// Specify storage ID parameter. Most devices that have 0x10001 have the storage ID. This
	// should be changed to use the device's real storage ID (which can be obtained by
	// removing the prefix for the WPD object ID for the storage)
	if (params != NULL && param_count > 0)
	{
		for (int i = 0; i < param_count; i++)
		{
			pvParam.ulVal = params[i];
			hr = spMtpParams->Add(&pvParam);
		}
	}

	// Add MTP parameters collection to our main parameter list
	if (hr == S_OK)
	{
		hr = spParameters->SetIPortableDevicePropVariantCollectionValue(
			WPD_PROPERTY_MTP_EXT_OPERATION_PARAMS, spMtpParams.Get());
	}

	// Send the command to the MTP device
	ComPtr<IPortableDeviceValues> spResults;
	if (hr == S_OK)
	{
		hr = pDevice->SendCommand(0, spParameters.Get(), &spResults);
	}

	// Check if the driver was able to send the command by interrogating WPD_PROPERTY_COMMON_HRESULT
	HRESULT hrCmd = S_OK;
	if (hr == S_OK)
	{
		hr = spResults->GetErrorValue(WPD_PROPERTY_COMMON_HRESULT, &hrCmd);
	}

	if (hr == S_OK)
	{
		//printf("Driver return code: 0x%08X\n", hrCmd);
		hr = hrCmd;
	}

	// If the command was executed successfully, check the MTP response code to see if the
	// device can handle the command. Be aware that there is a distinction between the command
	// being successfully sent to the device and the command being handled successfully by the device.
	if (hr == S_OK)
	{
		hr = spResults->GetUnsignedIntegerValue(WPD_PROPERTY_MTP_EXT_RESPONSE_CODE, result);
	}

	if (hr == S_OK)
	{
		//printf("MTP Response code: 0x%X\n", *result);
		hr = (*result == (DWORD)PTP_RESPONSECODE_OK) ? S_OK : E_FAIL;
	}

	// If the command was executed successfully, the MTP response parameters are returned in 
	// the WPD_PROPERTY_MTP_EXT_RESPONSE_PARAMS property, which is a PropVariantCollection
	ComPtr<IPortableDevicePropVariantCollection> spRespParams;
	if (hr == S_OK)
	{
		hr = spResults->GetIPortableDevicePropVariantCollectionValue(WPD_PROPERTY_MTP_EXT_RESPONSE_PARAMS,
			&spRespParams);
	}

	return hr;
}