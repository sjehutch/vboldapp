// =DOC FILE Interface for the DValidateAreaMonPtInput class
// =DOC TEXT Declares DValidateAreaMonPtInput class to validate input values for a monitor point
//////////////////////////////////////////////////////////////////

// ValidateAreaMonPtInput.h: interface for the DValidateAreaMonPtInput class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VALIDATEAREAMONPTINPUT_H__BAEE48A7_C6F1_4E1A_B359_8C7785F46136__INCLUDED_)
#define AFX_VALIDATEAREAMONPTINPUT_H__BAEE48A7_C6F1_4E1A_B359_8C7785F46136__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "AreaMonPtData.h"


// =DOC SECTION Class DValidateAreaMonPtInput
// =DOC SECTION Validate the input properties of an Area Monitor Point
// =DOC TEXT class DValidateAreaMonPtInput extends nothing
// =DOC TEXT Validate the input properties of an Area Monitor Point
//////////////////////////////////////////////////////////////////
class DValidateAreaMonPtInput  
{
public:
	// =DOC SECTION Method ValidateAreaMonPt( LPAREA_MON_PT_DATA lpAreaMonPtData ): BOOL
	// =DOC SECTION Validate the data supplied for an Area Monitor Point
	// =DOC TEXT public ValidateAreaMonPt( LPAREA_MON_PT_DATA lpAreaMonPtData ): BOOL
	// =DOC TEXT LPAREA_MON_PT_DATA lpAreaMonPtData: A carrier structure for the monitor point data
	// =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Validate the data supplied for an Area Monitor Point
	// =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateAreaMonPt( LPAREA_MON_PT_DATA lpAreaMonPtData );

   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DValidateAreaMonPtInput();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
	virtual ~DValidateAreaMonPtInput();

protected:
	// =DOC SECTION Method ValidateAchievedConcentrationTime( double &dConcentrationTime, CString &strError ): BOOL
	// =DOC SECTION Validate the ConcentrationTime data supplied for an Area Monitor Point
	// =DOC TEXT public ValidateAchievedConcentrationTime( double &dConcentrationTime, CString &strError ): BOOL
	// =DOC TEXT double &dConcentrationTime: ConcentrationTime data
   // =DOC TEXT CString &strError: String for error messages
	// =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Validate the ConcentrationTime data supplied for an Area Monitor Point
	// =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
	BOOL ValidateAchievedConcentrationTime( double &dConcentrationTime, CString &strError );

	// =DOC SECTION Method ValidateAchievedDose( double &dConcentrationTime, CString &strError ): BOOL
	// =DOC SECTION Validate the Dose data supplied for an Area Monitor Point
	// =DOC TEXT public ValidateAchievedDose( double &dConcentrationTime, CString &strError ): BOOL
	// =DOC TEXT double &dDose: Dose data
   // =DOC TEXT CString &strError: String for error messages
	// =DOC TEXT Returns BOOL: TRUE/FALSE
   // =DOC TEXT Validate the Dose data supplied for an Area Monitor Point
	// =DOC TEXT Return TRUE if there were no problems and FALSE if there were problems during the validation.
	//////////////////////////////////////////////////////////////////
   BOOL ValidateAchievedDose( double &dDose, CString &strError );
};

#endif // !defined(AFX_VALIDATEAREAMONPTINPUT_H__BAEE48A7_C6F1_4E1A_B359_8C7785F46136__INCLUDED_)
