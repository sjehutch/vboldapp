// =DOC FILE Validate concentration input values
// =DOC TEXT Declares DValidateConcentration class and tagConcentrationInputData structure
//////////////////////////////////////////////////////////////////

// ValidateConcentration.h: interface for the DValidateConcentration class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VALIDATECONCENTRATION_H__B3F8F477_3FC1_4775_8EB0_9C193ADC8241__INCLUDED_)
#define AFX_VALIDATECONCENTRATION_H__B3F8F477_3FC1_4775_8EB0_9C193ADC8241__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


// =DOC SECTION Struct tagConcentrationInputData ( CONC_INPUT_DATA, *LPCONC_INPUT_DATA )
// =DOC SECTION Carrier structure containing concentration information
// =DOC TEXT struct tagConcentrationInputData extends nothing
// =DOC TEXT Carrier structure containing concentration information
//////////////////////////////////////////////////////////////////
typedef struct tagConcentrationInputData
{
   // =DOC SECTION Member dConcentration: double
   // =DOC SECTION Concentration (g/m^3)
   // =DOC TEXT dConcentration: double
   // =DOC TEXT Concentration (g/m^3)
   //////////////////////////////////////////////////////////////////
   double dConcentration;

   // =DOC SECTION Member odtDateTime: COleDateTime
   // =DOC SECTION Date/time of the concentration
   // =DOC TEXT odtDateTime: COleDateTime
   // =DOC TEXT Date/time of the concentration
   //////////////////////////////////////////////////////////////////
   COleDateTime odtDateTime;

   // =DOC SECTION Member strMsg: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strMsg: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strMsg;

} CONC_INPUT_DATA, *LPCONC_INPUT_DATA;


// =DOC SECTION Class DValidateConcentration
// =DOC SECTION Validate concentration data input
// =DOC TEXT class DValidateConcentration extends nothing
// =DOC TEXT Validate concentration data input.
//////////////////////////////////////////////////////////////////
class DValidateConcentration  
{
public:
	// =DOC SECTION Method ValidateConcentration( LPCONC_INPUT_DATA lpConcInputData ): BOOL
	// =DOC SECTION Validate concentration data input
	// =DOC TEXT public  ValidateConcentration( LPCONC_INPUT_DATA lpConcInputData ): BOOL
	// =DOC TEXT LPCONC_INPUT_DATA lpConcInputData: Concentration data
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate concentration data iput.
   // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateConcentration( LPCONC_INPUT_DATA lpConcInputData );
	
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
   DValidateConcentration();
	
   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DValidateConcentration();

protected:
	// =DOC SECTION Method ValidateDate( COleDateTime &odtDateTime, CString &strMsg ): BOOL
	// =DOC SECTION Validate the date/time
	// =DOC TEXT protected  ValidateDate( COleDateTime &odtDateTime, CString &strMsg ): BOOL
	// =DOC TEXT COleDateTime &odtDateTime: Date/time value
   // =DOC TEXT CString &strMsg: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate the date/time.
   // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateDate( const COleDateTime &odtDateTime, CString &strMsg );

	// =DOC SECTION Method ValidateConcentrationValue( double &dConcentration, CString &strMsg ): BOOL
	// =DOC SECTION Validate concentration
	// =DOC TEXT protected  ValidateConcentrationValue( double &dConcentration, CString &strMsg ): BOOL
	// =DOC TEXT double &dConcentration: Concentration value
   // =DOC TEXT CString &strMsg: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate concentration.
   // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateConcentrationValue( const double &dConcentration, CString &strMsg );
};

#endif // !defined(AFX_VALIDATECONCENTRATION_H__B3F8F477_3FC1_4775_8EB0_9C193ADC8241__INCLUDED_)
