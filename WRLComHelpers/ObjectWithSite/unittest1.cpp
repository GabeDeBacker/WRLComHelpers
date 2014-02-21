#include "stdafx.h"
#include "CppUnitTest.h"
#include "..\ObjectWithSiteImpl.h"
#include "..\ProfferServiceImpl.h"

using namespace Windows::Internal::WRL;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;
using namespace Microsoft::WRL;
using namespace std;

namespace ObjectWithSite
{
	class CAgileProfferServiceWithSite : public RuntimeClass <
		RuntimeClassFlags<RuntimeClassType::ClassicCom>,
		Implements <RuntimeClassFlags<RuntimeClassType::ClassicCom>,
		Windows::Internal::WRL::ObjectWithSite,
		AgileProfferService >>
	{
	};

	class CSimpleServiceProvider : public RuntimeClass <
		RuntimeClassFlags<RuntimeClassType::ClassicCom>,
		Implements<RuntimeClassFlags<RuntimeClassType::ClassicCom>,
		Windows::Internal::WRL::ObjectWithSite,
		ProfferService >>
	{
	public:
		HRESULT RuntimeClassInitialize(_In_ REFGUID guidService)
		{
			_guidService = guidService;
			_dwExpectedThread = GetCurrentThreadId();
			return S_OK;
		}

		HRESULT v_QueryService(_In_ REFGUID guidService, _In_ REFIID riid, _COM_Outptr_ void **ppv)
		{
			*ppv = nullptr;
			HRESULT hr = E_NOTIMPL;
			if (guidService == _guidService)
			{
				Assert::AreEqual(_dwExpectedThread, GetCurrentThreadId());
				hr = CastToUnknown()->QueryInterface(riid, ppv);
			}
			return hr;
		}

	private:
		GUID _guidService;
		DWORD _dwExpectedThread;
	};

	TEST_CLASS(TestObjectWithSite)
	{
	public:
		
		TEST_METHOD(TestCrossApartmentQueryService);
		TEST_METHOD(TestMultiLayerQueryService);
	};
	
	struct ObjectWithSiteTestData
	{
		IObjectWithSite *pTestObject;
		GUID guidService;
	};

	DWORD WINAPI _TestObjectWithSiteData(_In_ void *pObjectWithSiteTestData)
	{
		Assert::IsTrue(SUCCEEDED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)));
		auto pData = reinterpret_cast<ObjectWithSiteTestData *>(pObjectWithSiteTestData);
		ComPtr<IServiceProvider> spProvider;
		Assert::IsTrue(SUCCEEDED(pData->pTestObject->QueryInterface(IID_PPV_ARGS(&spProvider))));
		ComPtr<IServiceProvider> spTestQuery;
		Assert::IsTrue(SUCCEEDED(spProvider->QueryService(pData->guidService, IID_PPV_ARGS(&spTestQuery))));
		CoUninitialize();
		return 0;
	}
	
	void TestObjectWithSite::TestCrossApartmentQueryService()
	{
		GUID guid;
		Assert::IsTrue(SUCCEEDED(CoCreateGuid(&guid)));
		ComPtr<IServiceProvider> spSTAServiceProvider;
		Assert::IsTrue(SUCCEEDED(MakeAndInitialize<CSimpleServiceProvider>(&spSTAServiceProvider, guid)));
		ComPtr<IObjectWithSite> spFTMObjectWithSite;
		Assert::IsTrue(SUCCEEDED((MakeAndInitialize<CAgileProfferServiceWithSite>(&spFTMObjectWithSite))));
		Assert::IsTrue(SUCCEEDED(spFTMObjectWithSite->SetSite(spSTAServiceProvider.Get())));
		ObjectWithSiteTestData data = { spFTMObjectWithSite.Get(), guid };
		HANDLE hThread = CreateThread(nullptr, 0, _TestObjectWithSiteData, &data, 0, nullptr);
		if (hThread != nullptr)
		{
			DWORD dwIndex;
			CoWaitForMultipleHandles(COWAIT_DISPATCH_CALLS | COWAIT_DISPATCH_WINDOW_MESSAGES, INFINITE, 1, &hThread, &dwIndex);
			CloseHandle(hThread);
		}
		spFTMObjectWithSite->SetSite(nullptr);
	}

	void TestObjectWithSite::TestMultiLayerQueryService()
	{
		struct ServiceProviderInfo
		{
			ComPtr<IServiceProvider> spProvider;
			GUID serviceGUID;
		};
		vector<ServiceProviderInfo> rgProviders;
		ComPtr<IServiceProvider> spPreviousProvider;

		for (unsigned int idxProvider = 0; idxProvider < 10; idxProvider++)
		{
			ServiceProviderInfo info;
			Assert::IsTrue(SUCCEEDED(CoCreateGuid(&info.serviceGUID)));
			Assert::IsTrue(SUCCEEDED(MakeAndInitialize<CSimpleServiceProvider>(&info.spProvider, info.serviceGUID)));
			ComPtr<IObjectWithSite> spSite;
			Assert::IsTrue(SUCCEEDED(info.spProvider.CopyTo(IID_PPV_ARGS(&spSite))));
			Assert::IsTrue(SUCCEEDED(spSite->SetSite(spPreviousProvider.Get())));
			spPreviousProvider = info.spProvider;
			rgProviders.push_back(info);
		}

		for (auto info : rgProviders)
		{
			ComPtr<IServiceProvider> spProvider;
			Assert::IsTrue(SUCCEEDED(spPreviousProvider->QueryService(info.serviceGUID, IID_PPV_ARGS(&spProvider))));
		}

		for (auto info : rgProviders)
		{
			ComPtr<IObjectWithSite> spSite;
			Assert::IsTrue(SUCCEEDED(info.spProvider.CopyTo(IID_PPV_ARGS(&spSite))));
			Assert::IsTrue(SUCCEEDED(spSite->SetSite(nullptr)));
		}
	}

}