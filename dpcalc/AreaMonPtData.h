// =DOC FILE Structure for area monitor point information
// =DOC TEXT Declares tagAreaMonPtData structure
//////////////////////////////////////////////////////////////////

// AreaMonPtData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_AREAMONPTDATA_H_)
#define _AREAMONPTDATA_H_


#include <afxtempl.h>



// =DOC SECTION Struct tagAreaMonPtData ( AREA_MON_PT_DATA, *LPAREA_MON_PT_DATA )
// =DOC SECTION Carrier structure containing area monitor point information
// =DOC TEXT struct tagAreaMonPtData extends nothing
// =DOC TEXT Carrier structure containing area monitor point information
//////////////////////////////////////////////////////////////////
typedef struct tagAreaMonPtData
{
   // =DOC SECTION Member dAchievedConcentrationTime: double
   // =DOC SECTION Achieved concentration-time (g-hr/m^3)
   // =DOC TEXT dAchievedConcentrationTime: double
   // =DOC TEXT Achieved concentration-time (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dAchievedConcentrationTime;

   // =DOC SECTION Member dAchievedDose: double
   // =DOC SECTION Achieved dose (g-hr/m^3)
   // =DOC TEXT dAchievedDose: double
   // =DOC TEXT Achieved dose (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dAchievedDose;

   // =DOC SECTION Member dHlt: double
   // =DOC SECTION Half-loss-time, HLT (hrs)
   // =DOC TEXT dHlt: double
   // =DOC TEXT Half-loss-time, HLT (hrs)
   //////////////////////////////////////////////////////////////////
   double dHlt;

   // =DOC SECTION Member dProjectedConcentrationTime: double
   // =DOC SECTION Projected concentration-time (g-hr/m^3)
   // =DOC TEXT dProjectedConcentrationTime: double
   // =DOC TEXT Projected concentration-time (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dProjectedConcentrationTime;

   // =DOC SECTION Member dProjectedTime: double
   // =DOC SECTION Projected for the fumigation (hrs)
   // =DOC TEXT dProjectedTime: double
   // =DOC TEXT Projected for the fumigation (hrs)
   //////////////////////////////////////////////////////////////////
   double dProjectedTime;

   // =DOC SECTION Member dFutureTargetConcentration: double
   // =DOC SECTION The future target concentration 1 hr after the current time (g/m^3)
   // =DOC TEXT dFutureTargetConcentration: double
   // =DOC TEXT The future target concentration 1 hr after the current time (g/m^3)
   //////////////////////////////////////////////////////////////////
   double dFutureTargetConcentration;
   
   // =DOC SECTION Member dFutureConcentration: double
   // =DOC SECTION The future concentration 1 hr after the current time (g/m^3)
   // =DOC TEXT dFutureConcentration: double
   // =DOC TEXT The future concentration 1 hr after the current time (g/m^3)
   //////////////////////////////////////////////////////////////////
   double dFutureConcentration;

   // =DOC SECTION Member dElapsedTime: double
   // =DOC SECTION Elapsed time of last concentration reading from start of exposure (hrs)
   // =DOC TEXT dElapsedTime: double
   // =DOC TEXT Elapsed time of last concentration reading from start of exposure (hrs)
   //////////////////////////////////////////////////////////////////
   double dElapsedTime;
   
   // =DOC SECTION Member strMsg: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strMsg: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strMsg;

} AREA_MON_PT_DATA, *LPAREA_MON_PT_DATA;


// =DOC SECTION CArray AREA_MON_PT_DATA_ARRAY
// =DOC SECTION Array of AREA_MON_PT_DATA structures
// =DOC TEXT CArray AREA_MON_PT_DATA_ARRAY extends CArray
// =DOC TEXT Array of AREA_MON_PT_DATA structures
//////////////////////////////////////////////////////////////////
typedef CArray < AREA_MON_PT_DATA, AREA_MON_PT_DATA& > AREA_MON_PT_DATA_ARRAY;


#endif // !defined(_AREAMONPTDATA_H_)
