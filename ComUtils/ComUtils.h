//****************************************************************************
// ComUtils.h
//****************************************************************************
// Helper functions for dealing with COM initialization and return values
// in a safe and efficient way.
//****************************************************************************
// (C) 2018 by Markus M. Egger https://markusegger.at drawing heavily
// from Kenny Kerr's (https://kennykerr.ca) Pluralsight Course on COM
// Essentials.
//****************************************************************************
// v0.6.1
// 2018-01-27T13:53:00+01
//****************************************************************************
// v0.6.1: Changed the return type of CheckHR to be more useful in corner
//         cases (conditionals) and added CheckHR_OKorFALSE to help
//         for example with COM enumerators like IEnumFORMATETC.
// v0.6.0: OLE Runtime wrapper added and tried to delete default constructors
//         for best practices/resilience.
// v0.5.0: Initial Release
//****************************************************************************

#pragma once


#include <tchar.h>
#include <wrl.h>


#ifndef TRACE
#define TRACE OutputDebugString
#endif


namespace Markus_M_Egger
{
	namespace ComUtils
	{
		// COM Apartment Types
		enum class Apartment
		{
			MultiThreaded = COINIT_MULTITHREADED,
			SingleThreaded = COINIT_APARTMENTTHREADED,
		};


		// COM function result checker prototypes
		inline bool CheckHR(const HRESULT& hr);
		inline bool CheckHR_OKorFALSE(const HRESULT& hr);


		// Type for COM exceptions
		struct ComException
		{
			explicit ComException(const HRESULT& hr) : result(hr)
			{
			}

			HRESULT hr(void) const
			{
				return result;
			}

		private:
			HRESULT result;
		};


		// Wrapper for COM runtime RAII
		struct ComRuntime
		{
			explicit ComRuntime(Apartment apartment)
			{
				initResult = CoInitializeEx(
					nullptr,
					static_cast<DWORD>(apartment)
				);

				CheckHR(initResult);

				TRACE(_T("COM runtime initialized."));
			}

			ComRuntime(const ComRuntime& other) = delete;
			ComRuntime(ComRuntime&& other) = delete;

			~ComRuntime(void)
			{
				if (S_OK == initResult)
				{
					CoUninitialize();

					TRACE(_T("COM runtime uninitialized."));
				}
			}

		private:
			HRESULT initResult{ E_FAIL };
		};


		// Wrapper for OLE COM runtime RAII
		struct OleRuntime
		{
			explicit OleRuntime(void)
			{
				initResult = OleInitialize(nullptr);
				
				CheckHR(initResult);

				TRACE(_T("OLE COM runtime initialized."));
			}

			OleRuntime(const OleRuntime& other) = delete;
			OleRuntime(OleRuntime&& other) = delete;

			~OleRuntime(void)
			{
				// "... each successful call to OleInitialize, including those
				// that return S_FALSE, must be balanced by a corresponding
				// call to OleUninitialize."
				if (S_OK == initResult
					|| S_FALSE == initResult)
				{
					OleUninitialize();

					TRACE(_T("OLE COM runtime uninitialized."));
				}
			}

		private:
			HRESULT initResult{ E_FAIL };
		};


		// COM function result checker
		inline bool CheckHR(const HRESULT& hr)
		{
			_ASSERTE(S_OK == hr);

			if (S_OK != hr)
			{
				throw ComException{ hr };
			}

			return true;
		}


		// COM function result checker - S_FALSE is ok, too.
		inline bool CheckHR_OKorFALSE(const HRESULT& hr)
		{
			_ASSERTE(S_OK == hr || S_FALSE == hr);

			if (!(S_OK == hr || S_FALSE == hr))
			{
				throw ComException{ hr };
			}

			return true;
		}
	}
}

//****************************************************************************
