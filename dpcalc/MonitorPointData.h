// =DOC FILE Structure for monitor point information
// =DOC TEXT Declares tagMonitorPointData structure
//////////////////////////////////////////////////////////////////

// MonitorPointStatusData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_MONITORPOINTDATA_H_)
#define _MONITORPOINTDATA_H_


#include <afxtempl.h>
#include "ConcentrationData.h"
#include "PestData.h"



// =DOC SECTION Struct tagMonitorPointStatusData ( MONITOR_POINT_STATUS_DATA, *LPMONITOR_POINT_STATUS_DATA )
// =DOC SECTION Carrier structure containing monitoring point information, both input and calculated
// =DOC TEXT struct tagMonitorPointStatusData extends nothing
// =DOC TEXT Carrier structure containing monitoring point information, both input and calculated
//////////////////////////////////////////////////////////////////
typedef struct tagMonitorPointData
{
   // User generated values.

   // =DOC SECTION Member ePest: PEST
   // =DOC SECTION Enumerated pest type
   // =DOC TEXT ePest: PEST
   // =DOC TEXT Enumerated pest type
   //////////////////////////////////////////////////////////////////
   PEST ePest; 
   
   // =DOC SECTION Member eFumigationType: FUMIGATION_TYPE
   // =DOC SECTION Enumerated fumigation type
   // =DOC TEXT eFumigationType: FUMIGATION_TYPE
   // =DOC TEXT Enumerated fumigation type
   //////////////////////////////////////////////////////////////////
   FUMIGATION_TYPE eFumigationType; 

   // =DOC SECTION Member eLifeStage: LIFE_STAGE
   // =DOC SECTION Enumerated life stage
   // =DOC TEXT eLifeStage: LIFE_STAGE
   // =DOC TEXT Enumerated life stage
   //////////////////////////////////////////////////////////////////
   LIFE_STAGE eLifeStage;
   
   // =DOC SECTION Member eTreatmentType: TREATMENT_TYPE
   // =DOC SECTION Enumerated treatment type
   // =DOC TEXT eTreatmentType: TREATMENT_TYPE
   // =DOC TEXT Enumerated treatment type
   //////////////////////////////////////////////////////////////////
   TREATMENT_TYPE eTreatmentType;
   
   // =DOC SECTION Member dTemp: double
   // =DOC SECTION User entered temperature (C)
   // =DOC TEXT dTemp: double
   // =DOC TEXT User entered temperature (C)
   //////////////////////////////////////////////////////////////////
   double dTemp;

   // =DOC SECTION Member dTimeExposure: double
   // =DOC SECTION User entered exposure time duration (hrs)
   // =DOC TEXT dTimeExposure: double
   // =DOC TEXT User entered exposure time duration (hrs)
   //////////////////////////////////////////////////////////////////
   double dTimeExposure;
   
   // =DOC SECTION Member dVolume: double
   // =DOC SECTION User entered volume (m^3)
   // =DOC TEXT dVolume: double
   // =DOC TEXT User entered volume (m^3)
   //////////////////////////////////////////////////////////////////
   double dVolume;
   
   // =DOC SECTION Member dHltOriginal: double
   // =DOC SECTION User entered HLT (hrs)
   // =DOC TEXT dHltOriginal: double
   // =DOC TEXT User entered HLT (hrs)
   //////////////////////////////////////////////////////////////////
   double dHltOriginal;

   // =DOC SECTION Member padataConc: CONCENTRATION_DATA_ARRAY*
   // =DOC SECTION Array of concentration data
   // =DOC TEXT padataConc: CONCENTRATION_DATA_ARRAY*
   // =DOC TEXT Array of concentration data
   //////////////////////////////////////////////////////////////////
   CONCENTRATION_DATA_ARRAY* padataConc;

   // Calculated values.

   // =DOC SECTION Member dAchievedDose: double
   // =DOC SECTION Calculated achieved dose
   // =DOC TEXT dAchievedDose: double
   // =DOC TEXT Calculated achieved dose
   //////////////////////////////////////////////////////////////////
   double dAchievedDose;

   // =DOC SECTION Member dAchievedCT: double
   // =DOC SECTION Calculated achieved CT
   // =DOC TEXT dAchievedCT: double
   // =DOC TEXT Calculated achieved CT
   //////////////////////////////////////////////////////////////////
   double dAchievedCT;

   // =DOC SECTION Member odtLastPeakDateTime: COleDateTime
   // =DOC SECTION Date/time for the last peak concentration
   // =DOC TEXT odtLastPeakDateTime: COleDateTime
   // =DOC TEXT Date/time for the last peak concentration
   //////////////////////////////////////////////////////////////////
   COleDateTime odtLastPeakDateTime;

   // =DOC SECTION Member odtHltDateTime: COleDateTime
   // =DOC SECTION Date/time for the concentration used to calculate last peak HLT
   // =DOC TEXT odtHltDateTime: COleDateTime
   // =DOC TEXT Date/time for the concentration used to calculate last peak HLT
   //////////////////////////////////////////////////////////////////
   COleDateTime odtHltDateTime;
   
   // =DOC SECTION Member odtsLastCTime: COleDateTimeSpan
   // =DOC SECTION Date/time span for the last concentration time
   // =DOC TEXT odtsLastCTime: COleDateTimeSpan
   // =DOC TEXT Date/time span for the last concentration time
   //////////////////////////////////////////////////////////////////
   COleDateTimeSpan odtsLastCTime;

   // =DOC SECTION Member dActualOHLT: double
   // =DOC SECTION Calculated actual OHLT
   // =DOC TEXT dActualOHLT: double
   // =DOC TEXT Calculated actual OHLT
   //////////////////////////////////////////////////////////////////
   double dActualOHLT;

   // =DOC SECTION Member dActualEHLT: double
   // =DOC SECTION Calculated actual EHLT
   // =DOC TEXT dActualEHLT: double
   // =DOC TEXT Calculated actual EHLT
   //////////////////////////////////////////////////////////////////
   double dActualEHLT;

   // =DOC SECTION Member dRemainingCT: double
   // =DOC SECTION Calculated remaining concentration-time to accumulate
   // =DOC TEXT dRemainingCT: double
   // =DOC TEXT Calculated remaining concentration-time to accumulate
   //////////////////////////////////////////////////////////////////
   double dRemainingCT;

   // =DOC SECTION Member dProjectedCT: double
   // =DOC SECTION Calculated projected concentration-time
   // =DOC TEXT dProjectedCT: double
   // =DOC TEXT Calculated projected concentration-time
   //////////////////////////////////////////////////////////////////
   double dProjectedCT;

   // =DOC SECTION Member dProjectedFumigationTime: double
   // =DOC SECTION Calculated projected total fumigation time
   // =DOC TEXT dProjectedFumigationTime: double
   // =DOC TEXT Calculated projected total fumigation time
   //////////////////////////////////////////////////////////////////
   double dProjectedFumigationTime;

   // =DOC SECTION Member dNewTargetC: double
   // =DOC SECTION Calculated new target concentration for 1 hour in the future
   // =DOC TEXT dNewTargetC: double
   // =DOC TEXT Calculated new target concentration for 1 hour in the future
   //////////////////////////////////////////////////////////////////
   double dNewTargetC;

   // =DOC SECTION Member dFutureC: double
   // =DOC SECTION Calculated future concentration
   // =DOC TEXT dFutureC: double
   // =DOC TEXT Calculated future concentration
   //////////////////////////////////////////////////////////////////
   double dFutureC;

   // =DOC SECTION Member dFutureDose: double
   // =DOC SECTION Calculated future dose
   // =DOC TEXT dFutureDose: double
   // =DOC TEXT Calculated future dose
   //////////////////////////////////////////////////////////////////
   double dFutureDose;

   // =DOC SECTION Member dRemainingDose: double
   // =DOC SECTION Calculated remaining dose required
   // =DOC TEXT dRemainingDose: double
   // =DOC TEXT Calculated remaining dose required
   //////////////////////////////////////////////////////////////////
   double dRemainingDose;

   // =DOC SECTION Member strMsg: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strMsg: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strMsg;

} MONITOR_POINT_DATA, *LPMONITOR_POINT_DATA;

#endif // !defined(_MONITORPOINTDATA_H_)
