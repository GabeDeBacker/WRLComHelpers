//
//    Copyright (C) Microsoft.  All rights reserved.
//
#pragma once
#include <wrl.h>                             // ComPtr, Runtimeclass, etc.
#include <wrl\Wrappers\CoreWrappers.h>       // For SRWLock
#include <ObjIdlbase.h>                      // For IAgileReference
#include <ShObjIdl.h>                        // For IObjectWithSite

// Class usage:
// Derive from this class when your object wants to be "sited" to another object.
// Typically the "site" of an object is used to discover functonality that the site exposes.
// Often the "site" exposes the IServiceProvider interface that can be used to discover functionality.
//
// The services this class provides is an abstraction between sites that may be agile (free threaded)
// in the site "chain" versus those that are not.
//
// For example, imagine a chain of sited objects such as below: (Agile = Thread Safe, Non-Agile means non-thread safe)
//  
//  | Object 1  |    | Object 2  |     | Object 3 |
//  | Non-Agile | -> | Non-Agile | ->  | Agile    |
//
// Since multiple threads can run inside the code of Object 3 (since it is agile) if the code attempts to acquire
// functionality that exists in Object 2 by acquiring COM interfaces through it's site chain, it is important
// that the call from the thread running in Object 3 through the site chain is marshaled to the proper thread to
// run inside Object 2.
//

namespace Windows { namespace Internal { namespace WRL {
class ObjectWithSite :
    public Microsoft::WRL::Implements<Microsoft::WRL::RuntimeClassFlags<Microsoft::WRL::RuntimeClassType::ClassicCom>,
                IObjectWithSite>
{
public:
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
