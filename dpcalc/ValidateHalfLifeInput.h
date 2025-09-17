// =DOC FILE Validate half loss time (HLT) input values
// =DOC TEXT Declares DValidateHalfLifeInput class and tagHalfLifeInputData structure
//////////////////////////////////////////////////////////////////

// ValidateHalfLifeInput.h: interface for the DValidateHalfLifeInput class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VALIDATEHALFLIFEINPUT_H__B54706F6_267E_437C_AB8E_F90CCA00B9AF__INCLUDED_)
#define AFX_VALIDATEHALFLIFEINPUT_H__B54706F6_267E_437C_AB8E_F90CCA00B9AF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


// =DOC SECTION Struct tagHalfLifeInputData ( HALFLIFE_INPUT_DATA, *LPHALFLIFE_INPUT_DATA )
// =DOC SECTION Carrier structure containing half life (HLT) information
// =DOC TEXT struct tagHalfLifeInputData extends nothing
// =DOC TEXT Carrier structure containing half life (HLT) information
//////////////////////////////////////////////////////////////////
typedef struct tagHalfLifeInputData
{
   // =DOC SECTION Member dHlt: double
   // =DOC SECTION User entered half loss time
   // =DOC TEXT dHlt: double
   // =DOC TEXT User entered half loss time
   //////////////////////////////////////////////////////////////////
   double dHlt; 

   // =DOC SECTION Member strMsg: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strMsg: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strMsg;

} HALFLIFE_INPUT_DATA, *LPHALFLIFE_INPUT_DATA;



// =DOC SECTION Class DValidateHalfLifeInput
// =DOC SECTION Validate half loss time data
// =DOC TEXT class DValidateHalfLifeInput extends nothing
// =DOC TEXT Validate half loss time data.
//////////////////////////////////////////////////////////////////
class DValidateHalfLifeInput  
{
public:
	// =DOC SECTION Method ValidateHalfLifeInput( LPHALFLIFE_INPUT_DATA lpdataHalfLifeInput ): BOOL
	// =DOC SECTION Validate half loss time data input
	// =DOC TEXT public  ValidateHalfLifeInput( LPHALFLIFE_INPUT_DATA lpdataHalfLifeInput ): BOOL
	// =DOC TEXT LPHALFLIFE_INPUT_DATA lpdataHalfLifeInput: Half loss time (HLT) data
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate half loss time data input
   // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateHalfLifeInput( LPHALFLIFE_INPUT_DATA lpdataHalfLifeInput );
	
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
   DValidateHalfLifeInput();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DValidateHalfLifeInput();

protected:
	// =DOC SECTION Method ValidateHlt( double& dHlt, CString &strMsg ): BOOL
	// =DOC SECTION Validate half loss time (HLT)
	// =DOC TEXT protected  ValidateHlt( double& dHlt, CString &strMsg ): BOOL
	// =DOC TEXT double& dHlt: Half loss time
	// =DOC TEXT CString &strMsg: Error message
	// =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Validate half loss time (HLT)
   // =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateHlt( double& dHlt, CString &strMsg );
};

#endif // !defined(AFX_VALIDATEHALFLIFEINPUT_H__B54706F6_267E_437C_AB8E_F90CCA00B9AF__INCLUDED_)
