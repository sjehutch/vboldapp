// =DOC FILE Validate dose input values
// =DOC TEXT Declares DValidateDoseInput class and tagDoseInputData structure
//////////////////////////////////////////////////////////////////

// ValidateDoseInput.h: interface for the DValidateDoseInput class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VALIDATEDOSEINPUT_H__75FE4BFD_5573_4C85_BA7D_28BABDFD6CD6__INCLUDED_)
#define AFX_VALIDATEDOSEINPUT_H__75FE4BFD_5573_4C85_BA7D_28BABDFD6CD6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


// =DOC SECTION Struct tagDoseInputData ( DOSE_INPUT_DATA, *LPDOSE_INPUT_DATA )
// =DOC SECTION Carrier structure containing dose input information
// =DOC TEXT struct tagDoseInputData extends nothing
// =DOC TEXT Carrier structure containing dose input information
//////////////////////////////////////////////////////////////////
typedef struct tagDoseInputData
{
   // =DOC SECTION Member nPest: int
   // =DOC SECTION User entered pest number
   // =DOC TEXT nPest: int
   // =DOC TEXT User entered pest number
   //////////////////////////////////////////////////////////////////
   int nPest;

   // =DOC SECTION Member nFumigationType: int
   // =DOC SECTION User entered fumigation type number
   // =DOC TEXT nFumigationType: int
   // =DOC TEXT User entered fumigation type number
   //////////////////////////////////////////////////////////////////
   int nFumigationType;

   // =DOC SECTION Member nLifeStage: int
   // =DOC SECTION User entered life stage number
   // =DOC TEXT nLifeStage: int
   // =DOC TEXT User entered life stage number
   //////////////////////////////////////////////////////////////////
   int nLifeStage;

   // =DOC SECTION Member nTreatmentType: int
   // =DOC SECTION User entered treatment type number
   // =DOC TEXT nTreatmentType: int
   // =DOC TEXT User entered treatment type number
   //////////////////////////////////////////////////////////////////
   int nTreatmentType;

   // =DOC SECTION Member nCommodity: int
   // =DOC SECTION User entered commodity number
   // =DOC TEXT nCommodity: int
   // =DOC TEXT User entered commodity number
   //////////////////////////////////////////////////////////////////
   int nCommodity;

   // =DOC SECTION Member dHlt: double
   // =DOC SECTION HLT value (hrs)
   // =DOC TEXT dHlt: double
   // =DOC TEXT HLT value (hrs)
   //////////////////////////////////////////////////////////////////
   double dHlt;

   // =DOC SECTION Member dTemp: double
   // =DOC SECTION User entered temperature (C)
   // =DOC TEXT dTemp: double
   // =DOC TEXT User entered temperature (C)
   //////////////////////////////////////////////////////////////////
   double dTemp;

   // =DOC SECTION Member dTimeExposure: double
   // =DOC SECTION User entered exposure time (hrs)
   // =DOC TEXT dTimeExposure: double
   // =DOC TEXT User entered exposure time (hrs)
   //////////////////////////////////////////////////////////////////
   double dTimeExposure;

   // =DOC SECTION Member dVolume: double
   // =DOC SECTION User entered volume (m^3)
   // =DOC TEXT dVolume: double
   // =DOC TEXT User entered volume (m^3)
   //////////////////////////////////////////////////////////////////
   double dVolume;

   // =DOC SECTION Member dLoadFactor: double
   // =DOC SECTION User entered Load Factor
   // =DOC TEXT dLoadFactor: double
   // =DOC TEXT User entered Load Factor
   //////////////////////////////////////////////////////////////////
   double dLoadFactor;

   // =DOC SECTION Member intIsComplete: int
   // =DOC SECTION This will indicate if we need to compute for HLT Sorption
   // =DOC TEXT intIsComplete: int
   // =DOC TEXT This will indicate if we need to compute for HLT Sorption
   //////////////////////////////////////////////////////////////////
   int intIsComplete;

   // =DOC SECTION Member nCounter: int
   // =DOC SECTION This will indicate if we are moving to the next record
   // =DOC TEXT nCounter: int
   // =DOC TEXT This will indicate if we are moving to the next record
   //////////////////////////////////////////////////////////////////
   int intCounter;

   // =DOC SECTION Member strError: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strError: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strError;

} DOSE_INPUT_DATA, *LPDOSE_INPUT_DATA;


// =DOC SECTION Class DValidateDoseInput
// =DOC SECTION Validate dose input data
// =DOC TEXT class DValidateDoseInput extends nothing
// =DOC TEXT Validate dose input data
//////////////////////////////////////////////////////////////////
class DValidateDoseInput  
{
public:
	// =DOC SECTION Method ValidateDoseInput( LPDOSE_INPUT_DATA lpdataDoseInput ): BOOL
	// =DOC SECTION Validate dose input data
	// =DOC TEXT public  ValidateDoseInput( LPDOSE_INPUT_DATA lpdataDoseInput ): BOOL
	// =DOC TEXT LPDOSE_INPUT_DATA lpdataDoseInput: Dose input data
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate dose input data.
   // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateDoseInput( LPDOSE_INPUT_DATA lpdataDoseInput );
	
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
   DValidateDoseInput();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DValidateDoseInput();

protected:
	// =DOC SECTION Method ValidateVolume( double &dVolume, CString &strError ): BOOL
	// =DOC SECTION Validate the volume
	// =DOC TEXT protected  ValidateVolume( double &dVolume, CString &strError ): BOOL
	// =DOC TEXT double &dVolume: Volume (m^3)
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the volume.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateVolume( const double &dVolume, CString &strError );

	// =DOC SECTION Method ValidateExposureTime( double &dTimeExposure, CString &strError ): BOOL
	// =DOC SECTION Validate the exposure time
	// =DOC TEXT protected  ValidateExposureTime( double &dTimeExposure, CString &strError ): BOOL
	// =DOC TEXT double &dTimeExposure: Exposure time (hrs)
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT VValidate the exposure time.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateExposureTime( const double &dTimeExposure, CString &strError );

	// =DOC SECTION Method ValidateTemp( double &dTemp, CString &strError ): BOOL
	// =DOC SECTION Validate the temperature
	// =DOC TEXT protected  ValidateTemp( double &dTemp, CString &strError ): BOOL
	// =DOC TEXT double &dTemp: Temperature (C)
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the temperature.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateTemp( const double &dTemp, CString &strError );

	// =DOC SECTION Method ValidateHlt( double &dHlt, CString &strError ): BOOL
	// =DOC SECTION Validate the HLT
	// =DOC TEXT protected  ValidateHlt( double &dHlt, CString &strError ): BOOL
	// =DOC TEXT double &dHlt: Half loss time (hrs)
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the HLT.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateHlt( const double &dHlt, CString &strError );

	// =DOC SECTION Method ValidateTreatmentType( int &nTreatmentType, CString &strError ): BOOL
	// =DOC SECTION Validate the pest treatment type number
	// =DOC TEXT protected  ValidateTreatmentType( int &nTreatmentType, CString &strError ): BOOL
	// =DOC TEXT int &nTreatmentType: Treatment type number
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the treatment type number
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateTreatmentType( const int &nTreatmentType, CString &strError );

	// =DOC SECTION Method ValidateLifeStage( int &nLifeStage, CString &strError ): BOOL
	// =DOC SECTION Validate the pest life stage number
	// =DOC TEXT protected  ValidateLifeStage( int &nLifeStage, CString &strError ): BOOL
	// =DOC TEXT int &nLifeStage: Life stage number
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the volume.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateLifeStage( const int &nLifeStage, CString &strError );

	// =DOC SECTION Method ValidatePest( int &nPest, CString &strError ): BOOL
	// =DOC SECTION Validate the pest number
	// =DOC TEXT protected  ValidatePest( int &nPest, CString &strError ): BOOL
	// =DOC TEXT int &nPest: Pest number
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the pest number.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidatePest( const int &nPest, CString &strError );

	// =DOC SECTION Method ValidateFumigationType( const int &nFumigationType, CString &strError ): BOOL
	// =DOC SECTION Validate the fumigation type number
	// =DOC TEXT protected  ValidateFumigationType( const int &nFumigationType, CString &strError ): BOOL
	// =DOC TEXT int &nFumigationType: Fumigation type number
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the fumigation type number.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateFumigationType( const int &nFumigationType, CString &strError );

	// =DOC SECTION Method ValidateCommodity( const int &nCommodity, CString &strError ): BOOL
	// =DOC SECTION Validate the commodity number
	// =DOC TEXT protected  ValidateCommodity( const int &nCommodity, CString &strError ): BOOL
	// =DOC TEXT int &nCommodity: Commodity number
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the commodity number.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateCommodity( const int &nCommodity, CString &strError );

	//Pat 1/17/2006 - Added ValidateLoadFactor function

	// =DOC SECTION Method ValidateLoadFactor( const int &dLoadFactor, CString &strError ): BOOL
	// =DOC SECTION Validate the commodity number
	// =DOC TEXT protected  ValidateLoadFactor( const int &dLoadFactor, CString &strError ): BOOL
	// =DOC TEXT int &nCommodity: Commodity number
    // =DOC TEXT CString &strError: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the commodity number.
    // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateLoadFactor( const double &dLoadFactor, CString &strError );
};

#endif // !defined(AFX_VALIDATEDOSEINPUT_H__75FE4BFD_5573_4C85_BA7D_28BABDFD6CD6__INCLUDED_)
