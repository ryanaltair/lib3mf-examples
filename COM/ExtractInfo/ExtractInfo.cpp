/*++

© 2015 netfabb GmbH (Original Author)
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:
* Redistributions of source code must retain the above copyright
notice, this list of conditions and the following disclaimer.
* Redistributions in binary form must reproduce the above copyright
notice, this list of conditions and the following disclaimer in the
documentation and/or other materials provided with the distribution.
* Neither the name of Microsoft Corporation nor netfabb GmbH nor the
names of its contributors may be used to endorse or promote products
derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL MICROSOFT AND/OR NETFABB BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Abstract:

ExtractInfo.cpp : 3MF Read Example

--*/

#include <tchar.h>
#include <Windows.h>
#include <iostream>
#include <string>
#include <vector>
#include <atlbase.h>
#include <algorithm>

// COM Includes of 3MF Library
#include "NMR_COMFactory.h"

// Use NMR namespace for the interfaces
using namespace NMR;


HRESULT ShowObjectProperties(_In_ ILib3MFModelObjectResource * pObject)
{
	HRESULT hResult;
	DWORD nNeededChars;
	std::vector<wchar_t> pBuffer;
	std::wstring sName;
	std::wstring sPartNumber;

	// Retrieve Mesh Name Length
	hResult = pObject->GetName(nullptr, 0, &nNeededChars);
	if (hResult != S_OK)
		return hResult;

	// Retrieve Mesh Name
	if (nNeededChars > 0) {
		pBuffer.resize(nNeededChars + 1);
		hResult = pObject->GetName(&pBuffer[0], nNeededChars + 1, nullptr);
		pBuffer[nNeededChars] = 0;
		sName = std::wstring(&pBuffer[0]);
	}

	// Output Name in local codepage
	std::string sAnsiName(sName.begin(), sName.end());
	std::cout << "   Name:            \"" << sAnsiName << "\"" << std::endl;

	// Retrieve Mesh Part Number Length
	hResult = pObject->GetPartNumber(nullptr, 0, &nNeededChars);
	if (hResult != S_OK)
		return hResult;

	// Retrieve Mesh Name
	if (nNeededChars > 0) {
		pBuffer.resize(nNeededChars + 1);
		hResult = pObject->GetPartNumber(&pBuffer[0], nNeededChars + 1, nullptr);
		pBuffer[nNeededChars] = 0;
		sPartNumber = std::wstring(&pBuffer[0]);
	}

	// Output Part number in local codepage
	std::string sAnsiPartNumber(sPartNumber.begin(), sPartNumber.end());
	std::cout << "   Part number:     \"" << sAnsiPartNumber << "\"" << std::endl;

	// Output Object type
	DWORD ObjectType;
	hResult = pObject->GetType(&ObjectType);
	if (hResult != S_OK)
		return hResult;

	switch (ObjectType) {
	case MODELOBJECTTYPE_MODEL:
		std::cout << "   Object type:     model" << std::endl;
		break;
	case MODELOBJECTTYPE_SUPPORT:
		std::cout << "   Object type:     support" << std::endl;
		break;
	case MODELOBJECTTYPE_OTHER:
		std::cout << "   Object type:     other" << std::endl;
		break;
	default:
		std::cout << "   Object type: invalid" << std::endl;
		break;

	}


	return S_OK;
}


HRESULT ShowMeshObjectInformation(_In_ ILib3MFModelMeshObject * pMeshObject)
{
	HRESULT hResult;
	DWORD nVertexCount, nTriangleCount;

	hResult = ShowObjectProperties(pMeshObject);
	if (hResult != S_OK)
		return hResult;

	// Retrieve Mesh Vertex Count
	hResult = pMeshObject->GetVertexCount(&nVertexCount);
	if (hResult != S_OK)
		return hResult;

	// Retrieve Mesh Vertex Count
	hResult = pMeshObject->GetTriangleCount(&nTriangleCount);
	if (hResult != S_OK)
		return hResult;

	// Output data
	std::cout << "   Vertex count:    " << nVertexCount << std::endl;
	std::cout << "   Triangle count:  " << nTriangleCount << std::endl;

	return S_OK;
}

HRESULT ShowComponentsObjectInformation(_In_ ILib3MFModelComponentsObject * pComponentsObject)
{
	HRESULT hResult;
	hResult = ShowObjectProperties(pComponentsObject);
	if (hResult != S_OK)
		return hResult;

	// Retrieve Component
	DWORD nComponentCount;
	DWORD nIndex;
	hResult = pComponentsObject->GetComponentCount(&nComponentCount);
	if (hResult != S_OK)
		return hResult;

	std::cout << "   Component count:    " << nComponentCount << std::endl;

	for (nIndex = 0; nIndex < nComponentCount; nIndex++) {
		DWORD nResourceID;
		CComPtr<ILib3MFModelComponent> pComponent;
		hResult = pComponentsObject->GetComponent(nIndex, &pComponent);
		if (hResult != S_OK)
			return hResult;

		hResult = pComponent->GetObjectResourceID(&nResourceID);
		if (hResult != S_OK)
			return hResult;

		std::cout << "   Component " << nIndex << ":    Object ID:   " << nResourceID << std::endl;
		BOOL bHasTransform;
		hResult = pComponent->HasTransform(&bHasTransform);
		if (hResult != S_OK)
			return hResult;


		if (bHasTransform) {
			MODELTRANSFORM Transform;

			// Retrieve Transform
			hResult = pComponent->GetTransform(&Transform);
			if (hResult != S_OK)
				return hResult;

			std::cout << "                   Transformation:  [ " << Transform.m_fFields[0][0] << " " << Transform.m_fFields[0][1] << " " << Transform.m_fFields[0][2] << " " << Transform.m_fFields[0][3] << " ]" << std::endl;
			std::cout << "                                    [ " << Transform.m_fFields[1][0] << " " << Transform.m_fFields[1][1] << " " << Transform.m_fFields[1][2] << " " << Transform.m_fFields[1][3] << " ]" << std::endl;
			std::cout << "                                    [ " << Transform.m_fFields[2][0] << " " << Transform.m_fFields[2][1] << " " << Transform.m_fFields[2][2] << " " << Transform.m_fFields[2][3] << " ]" << std::endl;
		}
		else {
			std::cout << "                   Transformation:  none" << std::endl;

		}


	}

	return S_OK;
}

int _tmain(int argc, _TCHAR* argv[])
{
	// General Variables
	HRESULT hResult;
	DWORD nInterfaceVersion;
	BOOL pbHasNext;

	// Objects
	CComPtr<ILib3MFModel> pModel;
	CComPtr<ILib3MFModelFactory> pFactory;
	CComPtr<ILib3MFModelReader> p3MFReader;
	CComPtr<ILib3MFModelBuildItemIterator> pBuildItemIterator;
	CComPtr<ILib3MFModelResourceIterator> pResourceIterator;


	std::cout << "------------------------------------------------------------------" << std::endl;
	std::cout << "3MF Read example" << std::endl;
	std::cout << "------------------------------------------------------------------" << std::endl;

	// Initialize COM
	hResult = CoInitialize(nullptr);
	if (hResult != S_OK) {
		std::cout << "could not initialize COM: " << std::hex << hResult << std::endl;
		return -1;
	}

	// Create Factory Object
	hResult = NMRCreateModelFactory(&pFactory);
	if (hResult != S_OK) {
		std::cout << "could not get Model Factory: " << std::hex << hResult << std::endl;
		return -1;
	}


	// Check 3MF Library Version
	hResult = pFactory->GetInterfaceVersion(&nInterfaceVersion);
	if (hResult != S_OK) {
		std::cout << "could not get 3MF Library version: " << std::hex << hResult << std::endl;
		return -1;
	}

	if ((nInterfaceVersion != NMR_APIVERSION_INTERFACE)) {
		std::cout << "invalid 3MF Library version: " << NMR_APIVERSION_INTERFACE << std::endl;
		return -1;
	}

	// Create Model Instance
	hResult = pFactory->CreateModel(&pModel);
	if (hResult != S_OK) {
		std::cout << "could not create model: " << std::hex << hResult << std::endl;
		return -1;
	}

	// Parse Arguments
	if (argc != 2) {
		std::cout << "Usage: " << std::endl;
		std::cout << "ExtractInfo.exe model.3mf" << std::endl;
		return -1;
	}

	// Read 3MF file
	std::wstring sFilename(argv[1]);
	std::string sAnsiFilename(sFilename.begin(), sFilename.end());


	// Create Model Writer
	hResult = pModel->QueryReader("3mf", &p3MFReader);
	if (hResult != S_OK) {
		std::cout << "could not create model reader: " << std::hex << hResult << std::endl;
		return -1;
	}

	// Import Model from 3MF File
	std::cout << "reading " << sAnsiFilename << "..." << std::endl;
	hResult = p3MFReader->ReadFromFile(sFilename.c_str());
	if (hResult != S_OK) {
		std::cout << "could not read file: " << std::hex << hResult << std::endl;
		return -1;
	}

	// Iterate through all the Objects
	hResult = pModel->GetObjects(&pResourceIterator);
	if (hResult != S_OK) {
		std::cout << "could not get object: " << std::hex << hResult << std::endl;
		return -1;
	}

	hResult = pResourceIterator->MoveNext(&pbHasNext);
	if (hResult != S_OK) {
		std::cout << "could not get next object: " << std::hex << hResult << std::endl;
		return -1;
	}

	while (pbHasNext) {
		CComPtr<ILib3MFModelResource> pResource;
		ILib3MFModelMeshObject * pMeshObject;
		ILib3MFModelComponentsObject * pComponentsObject;
		ModelResourceID ResourceID;

		// get current resource
		hResult = pResourceIterator->GetCurrent(&pResource);
		if (hResult != S_OK) {
			std::cout << "could not get resource: " << std::hex << hResult << std::endl;
			return -1;
		}

		// get resource ID
		hResult = pResource->GetResourceID(&ResourceID);
		if (hResult != S_OK) {
			std::cout << "could not get resource id: " << std::hex << hResult << std::endl;
			return -1;
		}

		// Query mesh interface
		hResult = pResource->QueryInterface(IID_Lib3MF_ModelMeshObject, (void **)&pMeshObject);
		if (hResult == S_OK) {
			std::cout << "------------------------------------------------------" << std::endl;
			std::cout << "mesh object #" << ResourceID << ": " << std::endl;

			// Show Mesh Object Information
			hResult = ShowMeshObjectInformation(pMeshObject);
			if (hResult != S_OK)
				return -1;
		}

		// Query component interface
		hResult = pResource->QueryInterface(IID_Lib3MF_ModelComponentsObject, (void **)&pComponentsObject);
		if (hResult == S_OK) {
			std::cout << "------------------------------------------------------" << std::endl;
			std::cout << "component object #" << ResourceID << ": " << std::endl;

			// Show Component Object Information
			hResult = ShowComponentsObjectInformation(pComponentsObject);
			if (hResult != S_OK)
				return -1;
		}

		hResult = pResourceIterator->MoveNext(&pbHasNext);
		if (hResult != S_OK) {
			std::cout << "could not get next object: " << std::hex << hResult << std::endl;
			return -1;
		}
	}


	// Iterate through all the Build items
	hResult = pModel->GetBuildItems(&pBuildItemIterator);
	if (hResult != S_OK) {
		std::cout << "could not get build items: " << std::hex << hResult << std::endl;
		return -1;
	}

	hResult = pBuildItemIterator->MoveNext(&pbHasNext);
	if (hResult != S_OK) {
		std::cout << "could not get next build item: " << std::hex << hResult << std::endl;
		return -1;
	}

	while (pbHasNext) {

		DWORD ResourceID;
		MODELTRANSFORM Transform;
		CComPtr<ILib3MFModelBuildItem> pBuildItem;
		// Retrieve Build Item
		hResult = pBuildItemIterator->GetCurrent(&pBuildItem);
		if (hResult != S_OK) {
			std::cout << "could not get build item: " << std::hex << hResult << std::endl;
			return -1;
		}

		// Retrieve Resource
		CComPtr<ILib3MFModelObjectResource> pObjectResource;
		hResult = pBuildItem->GetObjectResource(&pObjectResource);
		if (hResult != S_OK) {
			std::cout << "could not get build item resource: " << std::hex << hResult << std::endl;
			return -1;
		}

		// Retrieve Resource ID
		hResult = pObjectResource->GetResourceID(&ResourceID);
		if (hResult != S_OK) {
			std::cout << "could not get resource id: " << std::hex << hResult << std::endl;
			return -1;
		}

		// Output
		std::cout << "------------------------------------------------------" << std::endl;
		std::cout << "Build item (Object #" << ResourceID << "): " << std::endl;

		// Check Object Transform
		BOOL bHasTransform;
		hResult = pBuildItem->HasObjectTransform(&bHasTransform);
		if (hResult != S_OK) {
			std::cout << "could not get check object transform: " << std::hex << hResult << std::endl;
			return -1;
		}

		if (bHasTransform) {
			// Retrieve Transform
			hResult = pBuildItem->GetObjectTransform(&Transform);
			if (hResult != S_OK) {
				std::cout << "could not get object transform: " << std::hex << hResult << std::endl;
				return -1;
			}

			std::cout << "   Transformation:  [ " << Transform.m_fFields[0][0] << " " << Transform.m_fFields[0][1] << " " << Transform.m_fFields[0][2] << " " << Transform.m_fFields[0][3] << " ]" << std::endl;
			std::cout << "                    [ " << Transform.m_fFields[1][0] << " " << Transform.m_fFields[1][1] << " " << Transform.m_fFields[1][2] << " " << Transform.m_fFields[1][3] << " ]" << std::endl;
			std::cout << "                    [ " << Transform.m_fFields[2][0] << " " << Transform.m_fFields[2][1] << " " << Transform.m_fFields[2][2] << " " << Transform.m_fFields[2][3] << " ]" << std::endl;
		}
		else {
			std::cout << "   Transformation:  none" << std::endl;

		}

		// Retrieve Mesh Part Number Length
		std::wstring sPartNumber;
		DWORD nNeededChars;
		hResult = pBuildItem->GetPartNumber(nullptr, 0, &nNeededChars);
		if (hResult != S_OK)
			return hResult;

		// Retrieve Mesh Name
		if (nNeededChars > 0) {
			std::vector<wchar_t> pBuffer;
			pBuffer.resize(nNeededChars + 1);
			hResult = pBuildItem->GetPartNumber(&pBuffer[0], nNeededChars + 1, nullptr);
			pBuffer[nNeededChars] = 0;
			sPartNumber = std::wstring(&pBuffer[0]);
		}

		// Output Part number in local codepage
		std::string sAnsiPartNumber(sPartNumber.begin(), sPartNumber.end());
		std::cout << "   Part number:     \"" << sAnsiPartNumber << "\"" << std::endl;

		hResult = pBuildItemIterator->MoveNext(&pbHasNext);
		if (hResult != S_OK) {
			std::cout << "could not get next build item: " << std::hex << hResult << std::endl;
			return -1;
		}
	}

	std::cout << "------------------------------------------------------" << std::endl;


	std::cout << "done" << std::endl;


	return 0;
}

