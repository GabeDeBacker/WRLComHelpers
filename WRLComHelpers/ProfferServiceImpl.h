#pragma once
#include <new>                              // For std::nothrow
#include <combaseapi.h>                     // For RoGetAgileReference
#include <wrl.h>                            // For Microsoft::WRL::ComPtr and friends
#include <wrl/wrappers/corewrappers.h>      // for SRWLock implementation

// This header file aids in the implementation of the IProfferService and helper for the 
// IQueryService and IServiceProvider interfaces.
//
// CProfferService or CAgileProfferService are the classes below that you should derive from.
//
// Choose CProfferSerivce in the case where you know calls to IProfferService and IQueryService will only
// occur from the same thread.
//
// If you choose to handle failures IServiceProvider::QueryService of C[Agile]ProfferService then override
// v_QueryService.
//
// Chose CAgileProfferService if your class that derives from WRL/FTMBase (is Agile)
//
// Should you want to control the AgileReferenceOptions of this class from Default to Delayed Marshaling
// then derive your own class from CProfferService or CAigleProfferService and specify the desired marshaling options.
//
// class CDelayedMarshalProfferService : public CProfferService
// {
//     CDelayedMarshalProfferService() : CProfferService(AgileReferenceOptions::AGILEREFERENCE_DELAYEDMARSHAL)
//     {
//     }
// };
//
//

namespace Windows { namespace Internal { namespace WRL {

namespace Details
{

class ProfferServiceNoLock
{
public:
    // Need to define a destructor so when this pattern is used
    // auto lock = srwLock.LockXXX
    // the compiler does not generate the warning
    // 
    // profferserviceimpl.h(58) : error C4189: 'lock' : local variable is initialized but not referenced
    // 
    // which is why returning something like (int 0) doesn't work.
    //
    ~ProfferServiceNoLock()
    {
    }

    ProfferServiceNoLock LockExclusive() { return (*this); }
    ProfferServiceNoLock LockShared() { return (*this); }
};

template <typename LockType>
class ProfferServiceBase : public Microsoft::WRL::Implements<Microsoft::WRL::RuntimeClassFlags<Microsoft::WRL::RuntimeClassType::ClassicCom>, 
                                                               IProfferService,
                                                               IServiceProvider>
{
public:
    // IProfferService
    IFACEMETHODIMP ProfferService(_In_ REFGUID rguidService, _In_ IServiceProvider *psp, _Out_ DWORD *pdwCookie)
    {
        *pdwCookie = 0;
        Microsoft::WRL::ComPtr<IAgileReference> spReference;
        HRESULT hr = RoGetAgileReference(_agileReferenceOption, __uuidof(psp), psp, &spReference);
        if (SUCCEEDED(hr))
        {
            auto lock = _srwLock.LockExclusive();
            auto pWalk = _pServiceItems;
            while (pWalk != nullptr)
            {
                if (pWalk->guidService == rguidService)
                {
                    hr = HRESULT_FROM_WIN32(ERROR_ALREADY_REGISTERED);
                    break;
                }
                pWalk = pWalk->pNext;
            }

            if (SUCCEEDED(hr))
            {
                ServiceItem *pItem = new (std::nothrow) ServiceItem(rguidService, _dwNextCookie, spReference.Get(), _pServiceItems);
                hr = (pItem != nullptr) ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    _pServiceItems = pItem;
                    *pdwCookie = _dwNextCookie++;
                }
            }
        }
        return hr;
    }

    IFACEMETHODIMP RevokeService(_In_ DWORD dwCookie)
    {
        HRESULT hr = E_INVALIDARG;
        // This is not declared in minimum scope so that the release occurs
        // outside of holding the lock
        Microsoft::WRL::ComPtr<IAgileReference> spReferenceRelease;
        {
            auto lock =_srwLock.LockExclusive();
            auto pWalk = _pServiceItems;
            auto pPrev = _pServiceItems;
            while (pWalk != nullptr)
            {
                if (pWalk->dwCookie == dwCookie)
                {
                    spReferenceRelease.Attach(pWalk->spServiceProviderAgileReference.Detach());
                    pPrev->pNext = pWalk->pNext;
                    _pServiceItems = (pWalk == _pServiceItems) ? _pServiceItems->pNext : _pServiceItems;

                    delete pWalk;
                    hr = S_OK;    
                    break;
                }
                pPrev = pWalk;
                pWalk = pWalk->pNext;
            }
        }
        return hr;
    }

    IFACEMETHODIMP QueryService(_In_ REFGUID guidService, _In_ REFIID riid, _COM_Outptr_ void **ppv)
    {
        *ppv = nullptr;
        HRESULT hr = E_NOTIMPL;
        // Make sure to resolve the reference outside of the lock
        Microsoft::WRL::ComPtr<IAgileReference> spProfiderReference;
        {
            auto lock =_srwLock.LockShared();
            auto pWalk = _pServiceItems;
            while (pWalk != nullptr)
            {
                if (pWalk->guidService == guidService)
                {
                    spProfiderReference = pWalk->spServiceProviderAgileReference;
                    hr = S_OK;
                    break;
                }
                pWalk = pWalk->pNext;
            }
        }

        if (SUCCEEDED(hr))
        {
            Microsoft::WRL::ComPtr<IServiceProvider> spProvider;
            hr = spProfiderReference->Resolve(IID_PPV_ARGS(&spProvider));
            if (SUCCEEDED(hr))
            {
                hr = spProvider->QueryService(guidService, riid, ppv);
            }
        }

        // In theory, this should be checking explicitly for E_NOTIMPL per guidelines
        // for implementing QueryService. However, this isn't always the case.
        // We can get here one of two ways. Either the service wasn't found in our list
        // or the spProvider->QueryService call above failed.
        if (FAILED(hr))
        {
            hr = v_QueryService(guidService, riid, ppv);
        }

        // Now, if all that fails and the object supports IObjectWithSite then
        // proceed up the site chain.
        if (FAILED(hr))
        {
            Microsoft::WRL::ComPtr<IObjectWithSite> spSite;
            if (SUCCEEDED(CastToUnknown()->QueryInterface(IID_PPV_ARGS(&spSite))))
            {
                Microsoft::WRL::ComPtr<IServiceProvider> spProvider;
                if (SUCCEEDED(spSite->GetSite(IID_IServiceProvider, &spProvider)))
                {
                    hr = spProvider->QueryService(guidService, riid, ppv);
                }
            }
        }
        return hr;
    }

protected:
    virtual HRESULT v_QueryService(_In_ REFGUID /*guidService*/, _In_ REFIID /*riid*/, _COM_Outptr_ void ** /*ppv*/)
    {
        return E_NOTIMPL;
    }

    ProfferServiceBase(AgileReferenceOptions agileReferenceOption) : _dwNextCookie(1) , 
                                                                     _agileReferenceOption(agileReferenceOption),
                                                                     _pServiceItems(nullptr)
    {
    }

private:
    struct ServiceItem
    {
        ServiceItem(REFGUID serviceId, DWORD cookie, IAgileReference *pReference, ServiceItem *pNextItem) : pNext(pNextItem),
                                                                                                            dwCookie(cookie),
                                                                                                            spServiceProviderAgileReference(pReference),
                                                                                                            guidService(serviceId) 
        {
        }

        REFGUID guidService;
        Microsoft::WRL::ComPtr<IAgileReference> spServiceProviderAgileReference;
        DWORD const dwCookie;
        ServiceItem *pNext;

    private:
        // Create a private assignment operator to avoid error C4512
        // This occurs due to the "const"-ness of dwCookie and REFGUID guidService
        ServiceItem & operator=( const ServiceItem & );
    };

    ServiceItem *_pServiceItems;
    DWORD        _dwNextCookie;    // unique cookie index for next proffered service
    LockType     _srwLock;
    AgileReferenceOptions _agileReferenceOption;
};
} 
//Details namespace

// Here is an example of CProfferService usage:
// class CNonAgileObject : public RuntimeClass<
//                                   RuntimeClassFlags<RuntimeClassType::ClassicCom>,
//                                   ProfferService>
// {
//      public:
//      // IServiceProvider
//      HRESULT v_QueryService(_In_ REFGUID serviceId, _In_ REFIID riid, _COM_Outptr_ void **ppvObject) override
//      {
//          *ppvObject = nullptr;
//          HRESULT hr = E_NOTIMPL;
//          if (serviceId == SID_CNonAgileObjectService1)
//          {
//              // Expose your service code here
//          }
//          if (serviceId == SID_CNonAgileObjectService2)
//          {
//              // Expose your service code here
//          }
//      
//          if (hr == E_NOTIMPL)
//          {
//              // Choose to delegate up the site chain if you have one
//              hr = IUnknown_QueryService(_spunkSite.Get(), serviceId, riid, ppvObject);
//          }
//          return hr;
//      }
// };
//

class ProfferService : public Windows::Internal::WRL::Details::ProfferServiceBase<Windows::Internal::WRL::Details::ProfferServiceNoLock>
{
public:
    ProfferService(AgileReferenceOptions agileReferenceOptions = AgileReferenceOptions::AGILEREFERENCE_DEFAULT) : ProfferServiceBase<Windows::Internal::WRL::Details::ProfferServiceNoLock>(agileReferenceOptions)
    {
    }
};

// If you have an agile object then the implementation would be similar to above
// class CAgileObject : public RuntimeClass<RuntimeClassFlags<RuntimeClassType::ClassicCom>,
//                                          FtmBase,
//                                          AgileProfferService>>
// {
//      public:
//      // IServiceProvider
//      HRESULT v_QueryService(_In_ REFGUID serviceId, _In_ REFIID riid, _COM_Outptr_ void **ppvObject) override
//      {
//          *ppvObject = nullptr;
//          HRESULT hr = E_NOTIMPL;
//          if (serviceId == SID_CNonAgileObjectService1)
//          {
//              // Expose your service code here
//          }
//          if (serviceId == SID_CNonAgileObjectService2)
//          {
//              // Expose your service code here
//          }
//      
//          if (hr == E_NOTIMPL)
//          {
//              // Choose to delegate up the site chain if you have one
//              hr = IUnknown_QueryService(_spunkSite.Get(), serviceId, riid, ppvObject);
//          }
//          return hr;
//      }
// };

class AgileProfferService : public Windows::Internal::WRL::Details::ProfferServiceBase<Microsoft::WRL::Wrappers::SRWLock>
{
public:
    AgileProfferService(AgileReferenceOptions agileReferenceOptions = AgileReferenceOptions::AGILEREFERENCE_DEFAULT) : Details::ProfferServiceBase<Microsoft::WRL::Wrappers::SRWLock>(agileReferenceOptions)
    {
    }
};
} //namespace Windows
} //namespace Internal
} //namespace WRL
