// ShootingLine.cpp : Implementation of CProFumeCalculator
#include "stdafx.h"
#include "Dpcalc.h"

#include "ShootingLine.h"
#include "ShootingRate.h"
#include "MaxShootingRate.h"
#include "GlobalEnums.h"


/////////////////////////////////////////////////////////////////////////////
// CShootingLine

//------------------------------------------------------------------------------
CShootingLine::CShootingLine()
{
   m_pMaxShootingRate = new DMaxShootingRate();
   m_pShootingRate = new DShootingRate();
   m_pValidator = new DValidateShootingLine();
}

//------------------------------------------------------------------------------
CShootingLine::~CShootingLine()
{
   if( m_pMaxShootingRate )
   {
      delete m_pMaxShootingRate;
      m_pMaxShootingRate = NULL;
   }
   
   if( m_pShootingRate )
   {
      delete m_pShootingRate;
      m_pShootingRate = NULL;
   }

   if( m_pValidator )
   {
      delete m_pValidator;
      m_pValidator = NULL;
   }
}

STDMETHODIMP CShootingLine::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IShootingLine
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
      if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

//------------------------------------------------------------------------------
STDMETHODIMP CShootingLine::GetShootingRate( double CylinderTemp,
                                             double OneFourthLength,
                                             double OneEightLength,
											 double NumCylinders,
											 double PressureChosen,
                                             double* ShootingRate,
											 double* TimeToEmptyCylinder,
                                             BSTR* ErrMsg )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   HRESULT hrRetVal = S_OK;

   SHOOTINGRATE_DATA tagData;
   tagData.dCylinderTemp = CylinderTemp;
   tagData.dOneFourthLength = OneFourthLength;
   tagData.dOneEightLength = OneEightLength;
   tagData.dPressureChosen = PressureChosen;
   tagData.dShootingRate = 0.0;
   tagData.dTimeToEmptyCylinder = 0.0; // Jett 3/14/2005 - init new variable
   tagData.dNumCylinders = NumCylinders;

   // If the shooting rate data is valid, calculate the shooting rate and 
   // validate the result
   if ( m_pValidator->ValidateShootingRateData( &tagData ) )
   {

      if( m_pShootingRate->GetShootingRate( &tagData ) )
      {
		 // SIMS #2103035 - Jett, return 0.1 if Shooting Rate is less than 0 
		 // negative shooting rates are not possible in the physical world
		 if (tagData.dShootingRate < 0)
			*ShootingRate = 0.1;
		 else
			*ShootingRate = tagData.dShootingRate;


         // Jett 3/15/2005 - return the time to empty Cylinder
		 *TimeToEmptyCylinder = tagData.dTimeToEmptyCylinder;

         // There may be warnings with the calc
         if( !tagData.strError.IsEmpty() )
         {
            *ErrMsg = tagData.strError.AllocSysString();
         }
      }
      
      // If the shooting rate is invalid (> max); set the rate to the max,
      // and return an error code and message
      else
      {
         *ErrMsg = tagData.strError.AllocSysString();

         hrRetVal = S_FALSE;
      }
   }
   
   // If the shooting rate data is invalid, return an error code and message
   else
   {
      hrRetVal = S_FALSE;
      *ErrMsg = tagData.strError.AllocSysString();
   }

	return hrRetVal;
}

//------------------------------------------------------------------------------
STDMETHODIMP CShootingLine::GetMaxShootingRate( double FanCapacity,
                                                int RelativeHumidity,
                                                double* MaxShootingRate,
                                                BSTR* ErrMsg )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   HRESULT hrRetVal = S_OK;

   MAXSHOOTINGRATE_DATA tagData;
   tagData.dFanCapacity = FanCapacity;
   tagData.nRelativeHumidity = RelativeHumidity;
   tagData.dMaxShootingRate = 0.0;

   // If the maximum shooting rate data is valid, then calculate the max shooting rate
   if ( m_pValidator->ValidateMaxShootingRateData( &tagData ) )
   {
      if( m_pMaxShootingRate->GetMaxShootingRate( &tagData ) )
      {
         *MaxShootingRate = tagData.dMaxShootingRate;

         // There may be warnings with the calc
         if( !tagData.strError.IsEmpty() )
         {
            *ErrMsg = tagData.strError.AllocSysString();
         }
      }
      else
      {
         *ErrMsg = tagData.strError.AllocSysString();

         hrRetVal = S_FALSE;
      }
   }
   
   // If the maximum shooting rate data is invalid, return an error code and message.
   else
   {
      hrRetVal = S_FALSE;
      *ErrMsg = tagData.strError.AllocSysString();
   }

	return hrRetVal;
}
