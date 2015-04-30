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

Converter.cpp : Can convert 3MFs to STL and back

--*/

#include <tchar.h>
#include <Windows.h>
#include <iostream>
#include <string>
#include <atlbase.h>
#include <algorithm>

// COM Includes of 3MF Library
#include "NMR_COMFactory.h"

// Use NMR namespace for the interfaces
using namespace NMR;

int _tmain(int argc, _TCHAR* argv[])
{
	// General Variables
	HRESULT hResult;
	ULONGLONG nStartTicks;
	DWORD nInterfaceVersion;

	// Objects
	CComPtr<ILib3MFModel> pModel;
	CComPtr<ILib3MFModelFactory> pFactory;
	CComPtr<ILib3MFModelReader> pReader;
	CComPtr<ILib3MFModelWriter> pWriter;


	std::cout << "------------------------------------------------------------------" << std::endl;
	std::cout << "3MF Model Converter" << std::endl;
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


	// Parse Arguments
	if (argc != 2) {
		std::cout << "Usage: " << std::endl;
		std::cout << "Convert 3MF to STL: Converter.exe model.3mf" << std::endl;
		std::cout << "Convert STL to 3MF: Converter.exe model.stl" << std::endl;
		return -1;
	}

	// Extract Extension of filename
	std::wstring sReaderName;
	std::wstring sWriterName;
	std::wstring sNewExtension;
	std::wstring sFilename(argv[1]);
	std::wstring sExtension = PathFindExtension(sFilename.c_str());
	std::transform(sExtension.begin(), sExtension.end(), sExtension.begin(), ::tolower);

	// Convert to Ansi
	std::string sAnsiFilename(sFilename.begin(), sFilename.end());
	std::string sAnsiExtension(sExtension.begin(), sExtension.end());

	// Which Reader and Writer classes do we need?
	if (sExtension == L".stl") {
		sReaderName = L"stl";
		sWriterName = L"3mf";
		sNewExtension = L".3mf";
	}
	if (sExtension == L".3mf") {
		sReaderName = L"3mf";
		sWriterName = L"stl";
		sNewExtension = L".stl";
	}
	if (sReaderName.length() == 0) {
		std::cout << "unknown input file extension:" << sAnsiExtension << std::endl;
		return -1;
	}

	// Create new filename
	std::wstring sOutputFilename = sFilename;
	sOutputFilename.erase(sOutputFilename.length() - sExtension.length());
	sOutputFilename += sNewExtension;
	std::string sAnsiOutputFilename(sOutputFilename.begin(), sOutputFilename.end());

	// Create Model Instance
	hResult = pFactory->CreateModel(&pModel);
	if (hResult != S_OK) {
		std::cout << "could not create model: " << std::hex << hResult << std::endl;
		return -1;
	}

	// Create Model Reader
	hResult = pModel->QueryReader(sReaderName.c_str(), &pReader);
	if (hResult != S_OK) {
		std::cout << "could not create model reader: " << std::hex << hResult << std::endl;
		return -1;
	}

	// Import Model from File
	std::cout << "reading " << sAnsiFilename << "..." << std::endl;
	nStartTicks = GetTickCount64 ();
	hResult = pReader->ReadFromFile(sFilename.c_str());
	if (hResult != S_OK) {
		std::cout << "could not parse file: " << std::hex << hResult << std::endl;
		return -1;
	}
	std::cout << "elapsed time: " << (GetTickCount64() - nStartTicks) << "ms" << std::endl;

	// Create Model Writer
	hResult = pModel->QueryWriter(sWriterName.c_str(), &pWriter);
	if (hResult != S_OK) {
		std::cout << "could not create model reader: " << std::hex << hResult << std::endl;
		return -1;
	}

	// Export Model into File
	std::cout << "writing " << sAnsiOutputFilename << "..." << std::endl;
	nStartTicks = GetTickCount64();
	hResult = pWriter->WriteToFile(sOutputFilename.c_str());
	if (hResult != S_OK) {
		std::cout << "could not write file: " << std::hex << hResult << std::endl;
		return -1;
	}
	std::cout << "elapsed time: " << (GetTickCount64() - nStartTicks) << "ms" << std::endl;
	std::cout << "done" << std::endl;

	return 0;
}

