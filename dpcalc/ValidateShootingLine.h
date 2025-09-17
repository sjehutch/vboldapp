// =DOC FILE Validate Shooting Line values
// =DOC TEXT Declares DValidateShootingLine class and tagShootingRateData structure

// ValidateShootingLine.h: interface for the DValidateShootingLine class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DVALIDATESHOOTINGLINE_H__7E9B6E73_049C_4111_B879_53CD95F31D76__INCLUDED_)
#define AFX_DVALIDATESHOOTINGLINE_H__7E9B6E73_049C_4111_B879_53CD95F31D76__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// =DOC SECTION Struct tagShootingRateData ( SHOOTINGRATE_DATA, *LPSHOOTINGRATE_DATA )
// =DOC SECTION Carrier structure containing Shooting Rate information
// =DOC TEXT struct tagShootingRateData extends nothing
// =DOC TEXT Carrier structure containing Shooting Rate information
//////////////////////////////////////////////////////////////////
typedef struct tagShootingRateData
{
   // =DOC SECTION Member dCylinderTemp: double
   // =DOC SECTION cylinder temperature
   // =DOC TEXT dCylinderTemp: double
   // =DOC TEXT cylinder temperature
   //////////////////////////////////////////////////////////////////
   double dCylinderTemp;

   // =DOC SECTION Member dHoseLength: double
   // =DOC SECTION hose length
   // =DOC TEXT dHoseLength: double
   // =DOC TEXT hose length
   //////////////////////////////////////////////////////////////////
   


   // =DOC SECTION Member dHoseDiameter: double
   // =DOC SECTION hose inside diameter
   // =DOC TEXT dHoseDiameter: double
   // =DOC TEXT hose inside diameter
   //////////////////////////////////////////////////////////////////

   //Pat 12/29/2004 - Removed these lines
   //double dHoseLength;

   // =DOC SECTION Member dHoseDiameter: double
   // =DOC SECTION hose inside diameter
   // =DOC TEXT dHoseDiameter: double
   // =DOC TEXT hose inside diameter
   //////////////////////////////////////////////////////////////////
   //double dHoseDiameter;

   // =DOC SECTION Member dShootingRate: double
   // =DOC SECTION shooting rate (calculated value)
   // =DOC TEXT dShootingRate: double
   // =DOC TEXT shooting rate (calculated value)
   //////////////////////////////////////////////////////////////////
   double dShootingRate;

   // =DOC SECTION Member strError: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strError: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////

   //Pat 12/29/2004 - Added these two variables	
   double dOneFourthLength;

   double dOneEightLength;

   double dPressureChosen;

   //Jett 1/14/2005 - added these two member variables
   double dTimeToEmptyCylinder;
   double dNumCylinders;
   
   CString strError;


	
} SHOOTINGRATE_DATA, *LPSHOOTINGRATE_DATA;



// =DOC SECTION Struct tagMaxShootingRateData ( MAXSHOOTINGRATE_DATA, *LPMAXSHOOTINGRATE_DATA )
// =DOC SECTION Carrier structure containing Maximum Shooting Rate information
// =DOC TEXT struct tagMaxShootingRateData extends nothing
// =DOC TEXT Carrier structure containing Maximum Shooting Rate information
//////////////////////////////////////////////////////////////////
typedef struct tagMaxShootingRateData
{
   // =DOC SECTION Member dFanCapacity: double
   // =DOC SECTION fan capacity
   // =DOC TEXT dFanCapacity: double
   // =DOC TEXT fan capacity
   //////////////////////////////////////////////////////////////////
   double dFanCapacity;

   // =DOC SECTION Member nRelativeHumidity: integer
   // =DOC SECTION relative humidity
   // =DOC TEXT nRelativeHumidity: integer
   // =DOC TEXT relative humidity
   //////////////////////////////////////////////////////////////////
   int nRelativeHumidity;

   // =DOC SECTION Member dMaxShootingRate: double
   // =DOC SECTION maximum shooting rate (calculated value)
   // =DOC TEXT dMaxShootingRate: double
   // =DOC TEXT maximum shooting rate (calculated value)
   //////////////////////////////////////////////////////////////////
   double dMaxShootingRate;

   // =DOC SECTION Member strError: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strError: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strError;

} MAXSHOOTINGRATE_DATA, *LPMAXSHOOTINGRATE_DATA;



// =DOC SECTION Class DValidateShootingLine
// =DOC SECTION Validate Shooting Line data
// =DOC TEXT class DValidateShootingLine extends nothing
// =DOC TEXT Validate Shooting Line data
//////////////////////////////////////////////////////////////////
class DValidateShootingLine  
{
public:

	// =DOC SECTION Method ValidateMaxShootingRateData( LPMAXSHOOTINGRATE_DATA lpdataMaxShootingRate ): BOOL
	// =DOC SECTION Validate Max Shooting Rate data
	// =DOC TEXT public ValidateMaxShootingRateData( LPMAXSHOOTINGRATE_DATA lpdataMaxShootingRate ): BOOL
	// =DOC TEXT lpdataMaxShootingRate: Max Shooting Rate data
	// =DOC TEXT Returns BOOL: success or failure
	// =DOC TEXT Validate Max Shooting Rate data
	//////////////////////////////////////////////////////////////////
	BOOL ValidateMaxShootingRateData( LPMAXSHOOTINGRATE_DATA lpdataMaxShootingRate );

	// =DOC SECTION Method ValidateShootingRate( LPSHOOTINGRATE_DATA lpdataShootingRate ): BOOL
	// =DOC SECTION Validate Shooting Rate data
	// =DOC TEXT public ValidateShootingRate( LPSHOOTINGRATE_DATA lpdataShootingRate ): BOOL
	// =DOC TEXT lpdataShootingRate: Shooting Rate data
	// =DOC TEXT Returns BOOL: success or failure
	// =DOC TEXT Validate Shooting Rate data
	//////////////////////////////////////////////////////////////////
	BOOL ValidateShootingRateData( LPSHOOTINGRATE_DATA lpdataShootingRate );

   // =DOC SECTION Method DValidateShootingLine(): constructor
	// =DOC SECTION Class constructor
	// =DOC TEXT public DValidateShootingLine()
	// =DOC TEXT Class constructor
	//////////////////////////////////////////////////////////////////
	DValidateShootingLine();

	// =DOC SECTION Method ~DValidateShootingLine(): destructor
	// =DOC SECTION Class destructor
	// =DOC TEXT public virtual ~DValidateShootingLine()
	// =DOC TEXT Class destructor
	//////////////////////////////////////////////////////////////////
   virtual ~DValidateShootingLine();

protected:
	// =DOC SECTION Method ValidateRelativeHumidity( const int &nRelativeHumidity, CString &strError ): BOOL
	// =DOC SECTION Validate the relative humidity
	// =DOC TEXT protected ValidateRelativeHumidity( const int &nRelativeHumidity, CString &strError ): BOOL
   // =DOC TEXT nRelativeHumidity: relative humidity
   // =DOC TEXT strError: error message if unsuccessful
	// =DOC TEXT Returns BOOL: success or failure
	// =DOC TEXT Validate the relative humidity
	//////////////////////////////////////////////////////////////////
	BOOL ValidateRelativeHumidity( const int &nRelativeHumidity, CString &strError );

	// =DOC SECTION Method ValidateFanCapacity( const double &dFanCapacity, CString &strError ): BOOL
	// =DOC SECTION Validate the fan capacity
	// =DOC TEXT protected ValidateFanCapacity( const double &dFanCapacity, CString &strError ): BOOL
   // =DOC TEXT dFanCapacity: fan capacity
   // =DOC TEXT strError: error message if unsuccessful
	// =DOC TEXT Returns BOOL: success or failure
	// =DOC TEXT Validate the fan capacity
	//////////////////////////////////////////////////////////////////
	BOOL ValidateFanCapacity( const double &dFanCapacity, CString &strError );

	// =DOC SECTION Method ValidateCylinderTemp( const double &dCylinderTemp, CString &strError ): BOOL
	// =DOC SECTION Validate the cylinder temperature
	// =DOC TEXT protected ValidateCylinderTemp( const double &dCylinderTemp, CString &strError ): BOOL
	// =DOC TEXT dCylinderTemp: cylinder temperature
    // =DOC TEXT strError: error message if unsuccessful
	// =DOC TEXT Returns BOOL: success or failure
	// =DOC TEXT Validate the cylinder temperature
	//////////////////////////////////////////////////////////////////
	BOOL ValidateCylinderTemp( const double &dCylinderTemp, const double &dPressureChosen, CString &strError );

	// =DOC SECTION Method ValidateHoseLength( const double &dHoseLength, CString &strError ): BOOL
	// =DOC SECTION Validate the hose length
	// =DOC TEXT protected ValidateHoseLength( const double &dHoseLength, CString &strError ): BOOL
    // =DOC TEXT dHoseLength: hose length
    // =DOC TEXT strError: error message if unsuccessful
	// =DOC TEXT Returns BOOL: success or failure
	// =DOC TEXT Validate the hose length
	//////////////////////////////////////////////////////////////////
	BOOL ValidateHoseLength( const double &dHoseLength, CString &strError );

	// =DOC SECTION Method ValidateHoseDiameter( const double &dHoseDiameter, CString &strError ): BOOL
	// =DOC SECTION Validate the hose inside diameter
	// =DOC TEXT protected ValidateHoseDiameter( const double &dHoseDiameter, CString &strError ): BOOL
    // =DOC TEXT dHoseDiameter: hose inside diameter
    // =DOC TEXT strError: error message if unsuccessful
	// =DOC TEXT Returns BOOL: success or failure
	// =DOC TEXT Validate the hose diameter
	//////////////////////////////////////////////////////////////////
	BOOL ValidateHoseDiameter( const double &dHoseDiameter, CString &strError );

	//Pat 12/29/2004
	BOOL ValidateOneFourthLength(const double &dOneFourthLength,CString &strError);

	BOOL ValidateOneEightLength(const double &dOneEightLength,CString &strError);
};

#endif // !defined(AFX_DVALIDATESHOOTINGLINE_H__7E9B6E73_049C_4111_B879_53CD95F31D76__INCLUDED_)
