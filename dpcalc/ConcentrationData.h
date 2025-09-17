// =DOC FILE Structure for concentration information
// =DOC TEXT Declares tagConcentrationData structure and an array of tagConcentrationData structures
//////////////////////////////////////////////////////////////////

// ConcentrationData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_CONCENTRATIONDATA_H_)
#define _CONCENTRATIONDATA_H_


#include <afxtempl.h>


// =DOC SECTION Struct tagConcentrationData ( CONCENTRATION_DATA, *LPCONCENTRATION_DATA )
// =DOC SECTION Carrier structure containing concentration information
// =DOC TEXT struct tagConcentrationData extends nothing
// =DOC TEXT Carrier structure containing concentration information
//////////////////////////////////////////////////////////////////
typedef struct tagConcentrationData
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

} CONCENTRATION_DATA, *LPCONCENTRATION_DATA;


// =DOC SECTION CArray CONCENTRATION_DATA_ARRAY
// =DOC SECTION Array of CONCENTRATION_DATA structures
// =DOC TEXT CArray CONCENTRATION_DATA_ARRAY extends CArray
// =DOC TEXT Array of CONCENTRATION_DATA structures
//////////////////////////////////////////////////////////////////
typedef CArray < CONCENTRATION_DATA, CONCENTRATION_DATA& > CONCENTRATION_DATA_ARRAY;

#endif // !defined(_CONCENTRATIONDATA_H_)
