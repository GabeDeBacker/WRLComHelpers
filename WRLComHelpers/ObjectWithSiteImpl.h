//
//    Copyright (C) Microsoft.  All rights reserved.
//
#pragma once
#include <wrl\implements.h>
#include <wrl\Wrappers\CoreWrappers.h>
#include <objbase.h>

namespace Windows { namespace Internal { namespace WRL {
class ObjectWithSite :
    public Microsoft::WRL::Implements<Microsoft::WRL::RuntimeClassFlags<Microsoft::WRL::RuntimeClassType::ClassicCom>,
                IObjectWithSite>
{
public:
    ObjectWithSite()
    {
    }

    IFACEMETHODIMP SetSite(_In_opt_ IUnknown *punkSite)
    {
        HRESULT hr;
        Microsoft::WRL::ComPtr<IAgileReference> spunkSiteOldReference;

        if (punkSite == nullptr)
        {
            // Let the release occur outside of the lock
            auto lock = srwLock.LockExclusive();
            spunkSiteOldReference.Attach(_spunkSiteReference.Detach());
            hr = S_OK;
        }
        else
        {
            // Acquire new reference outside of any locks
            Microsoft::WRL::ComPtr<IAgileReference> spunkNewSiteReference;
            hr = RoGetAgileReference(AgileReferenceOptions::AGILEREFERENCE_DEFAULT, __uuidof(punkSite), punkSite, &spunkNewSiteReference);
            if (SUCCEEDED(hr))
            {
                // Let the release occur outside of the lock
                auto lock = srwLock.LockExclusive();
                spunkSiteOldReference.Attach(_spunkSiteReference.Detach());
                _spunkSiteReference.Attach(spunkNewSiteReference.Detach());
            }
        }
        return hr;
    }

    IFACEMETHODIMP GetSite(_In_ REFIID riid, _COM_Outptr_ void **ppvSite)
    {
        *ppvSite = nullptr;
        Microsoft::WRL::ComPtr<IAgileReference> spunkSiteReference;
        
        // Resolve reference outside the lock
        {
            auto lock = srwLock.LockShared();
            spunkSiteReference = _spunkSiteReference;
        }

        return spunkSiteReference ? spunkSiteReference->Resolve(riid, ppvSite) : E_NOTIMPL;
    }

protected:
    // This bit of trickery allows for you to do _spunkSite.Get() on your derived class which in turn
    // calls the function above. Yes, it is possible for this function to return a nullptr, but in reality
    // your "naked" site pointer (_spunkSite) can also be nullptr at any time as well.
    __declspec(property(get = _GetSitePtr)) Microsoft::WRL::ComPtr<IUnknown> _spunkSite;

    Microsoft::WRL::ComPtr<IUnknown> _GetSitePtr()
    {
        Microsoft::WRL::ComPtr<IUnknown> spSite;
        GetSite(IID_PPV_ARGS(&spSite));
        return spSite;
    }
    
private:
    _Guarded_by_(srwLock)
    Microsoft::WRL::ComPtr<IAgileReference> _spunkSiteReference;
    Microsoft::WRL::Wrappers::SRWLock srwLock;
};
} // namespace Windows
} // namespace Internal
} // namespace WRL
