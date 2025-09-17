// ValidateShootingLine.cpp: implementation of the DValidateShootingLine class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ValidateShootingLine.h"
#include "resource.h"
#include "GlobalEnums.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DValidateShootingLine::DValidateShootingLine()
{

}

DValidateShootingLine::~DValidateShootingLine()
{

}

//------------------------------------------------------------------------------
BOOL DValidateShootingLine::ValidateCylinderTemp(const double &dCylinderTemp, 
												 const double &dPressureChosen,
                                                 CString &strError)
{
   BOOL bRetVal;

   // Cylinder temp must be within the valid range
   //Pat 4/8/05 - Check 
   if (dPressureChosen == 1)
   {
		if ( dCylinderTemp >= MIN_TEMP && dCylinderTemp <= MAX_TEMP )
		{
		  bRetVal = TRUE;
		}
		else
		{
		  bRetVal = FALSE;
		  strError.Format( IDS_INVALID_CYLINDER_TEMP, MIN_TEMP, MAX_TEMP );
		}
   }
   else
   {
		if ( dCylinderTemp >= MIN_PRES && dCylinderTemp <= MAX_PRES )
		{
		  bRetVal = TRUE;
		}
		else
		{
		  bRetVal = FALSE;
		  strError.Format( IDS_INVALID_CYLINDER_PRES, MIN_PRES, MAX_PRES );
		}
   }
   
   return bRetVal;
}

//------------------------------------------------------------------------------
BOOL DValidateShootingLine::ValidateHoseLength( const double &dHoseLength, 
                                                CString &strError )
{
   BOOL bRetVal;

   // Hose length must be within the valid range
   if ( dHoseLength >= MIN_LINE_LENGTH && dHoseLength <= MAX_LINE_LENGTH )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      strError.Format( IDS_INVALID_HOSE_LENGTH, MIN_LINE_LENGTH, MAX_LINE_LENGTH );
   }
   
   return bRetVal;
}

//------------------------------------------------------------------------------
BOOL DValidateShootingLine::ValidateHoseDiameter( const double &dHoseDiameter, 
                                                  CString &strError )
{
   BOOL bRetVal;

   // JNM 04/01/03 - Fumiguide Enhancements 2003 - Removed SHOOTINGLINE_DIAMETER_FOUR.
   // Hose diameter must be a valid value
   if ( dHoseDiameter == SHOOTINGLINE_DIAMETER_ONE   ||
        dHoseDiameter == SHOOTINGLINE_DIAMETER_TWO   ||
        dHoseDiameter == SHOOTINGLINE_DIAMETER_THREE )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      strError.Format( IDS_INVALID_HOSE_DIAMETER, 
                       SHOOTINGLINE_DIAMETER_ONE, 
                       SHOOTINGLINE_DIAMETER_TWO, 
                       SHOOTINGLINE_DIAMETER_THREE);
   }

   return bRetVal;
}

//------------------------------------------------------------------------------
BOOL DValidateShootingLine::ValidateShootingRateData( LPSHOOTINGRATE_DATA lpdataShootingRate )
{
   BOOL bRetVal;
   CString strError;

   // Shooting rate variables must be valid
   if ( ValidateCylinderTemp( lpdataShootingRate->dCylinderTemp, 
	                          lpdataShootingRate->dPressureChosen, 
							  strError ) && 
        ValidateOneFourthLength( lpdataShootingRate->dOneFourthLength, strError )     &&
        ValidateOneEightLength( lpdataShootingRate->dOneEightLength, strError ) )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      lpdataShootingRate->strError = strError;
   }

   return bRetVal;
}

//------------------------------------------------------------------------------
BOOL DValidateShootingLine::ValidateMaxShootingRateData( LPMAXSHOOTINGRATE_DATA lpdataMaxShootingRate)
{
   BOOL bRetVal;
   CString strError;

   // Don't check dMaxShootingRate, this is a calculated value

   // Maximum Shooting Rate variables must be valid
   if ( ValidateFanCapacity( lpdataMaxShootingRate->dFanCapacity, strError ) && 
        ValidateRelativeHumidity( lpdataMaxShootingRate->nRelativeHumidity, strError ) )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      lpdataMaxShootingRate->strError = strError;
   }

   return bRetVal;
}

//------------------------------------------------------------------------------
BOOL DValidateShootingLine::ValidateFanCapacity(const double &dFanCapacity, CString &strError)
{
   BOOL bRetVal;

   // Fan capacity must be within the valid range
   if ( dFanCapacity >= MIN_FAN && dFanCapacity <= MAX_FAN )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      strError.Format( IDS_INVALID_FAN_CAPACITY, MIN_FAN, MAX_FAN );
   }
   
   return bRetVal;
}

//------------------------------------------------------------------------------
BOOL DValidateShootingLine::ValidateRelativeHumidity(const int &nRelativeHumidity, CString &strError)
{
   BOOL bRetVal;

   // Relative humidity must be within the valid range
   if ( nRelativeHumidity >= MIN_HUMIDITY && nRelativeHumidity <= MAX_HUMIDITY )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      strError.Format( IDS_INVALID_RELATIVE_HUMIDITY, MIN_HUMIDITY, MAX_HUMIDITY );
   }
   
   return bRetVal;
}

//------------------------------------------------------------------------------
//Pat 12/29/2004
BOOL DValidateShootingLine::ValidateOneFourthLength(const double &dOneFourthLength,CString &strError)
{
	BOOL bRetVal;
	  if ( dOneFourthLength >= MIN_ONEFOURTH_LENGTH && dOneFourthLength <= MAX_ONEFOURTH_LENGTH )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      strError.Format( IDS_INVALID_ONEFOURTH_LENGTH, MIN_ONEFOURTH_LENGTH, MAX_ONEFOURTH_LENGTH );
   }
   
   return bRetVal;
}

//------------------------------------------------------------------------------
//Pat 12/29/2004
BOOL DValidateShootingLine::ValidateOneEightLength(const double &dOneEightLength,CString &strError)
{
	BOOL bRetVal;
	  if ( dOneEightLength >= MIN_ONEEIGHT_LENGTH && dOneEightLength <= MAX_ONEEIGHT_LENGTH )
   {
      bRetVal = TRUE;
   }
   else
   {
      bRetVal = FALSE;
      strError.Format( IDS_INVALID_ONEEIGHT_LENGTH, MIN_ONEEIGHT_LENGTH, MAX_ONEEIGHT_LENGTH );
   }
   
   return bRetVal;
}