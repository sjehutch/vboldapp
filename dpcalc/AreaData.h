// =DOC FILE Structure for area information
// =DOC TEXT Declares tagAreaData structure
//////////////////////////////////////////////////////////////////

// AreaData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_AREADATA_H_)
#define _AREADATA_H_


#include <afxtempl.h>

#include "AreaMonPtData.h"
#include "DasCommodityData.h"
#include "DasFumigationTypeData.h"
#include "DasLifeStageData.h"
#include "DasPestData.h"
#include "DasTreatmentTypeData.h"


// =DOC SECTION Struct tagAreaData ( AREA_DATA, *LPAREA_DATA )
// =DOC SECTION Carrier structure containing area information, both input and calculated
// =DOC TEXT struct tagAreaData extends nothing
// =DOC TEXT Carrier structure containing area information, both input and calculated
//////////////////////////////////////////////////////////////////
typedef struct tagAreaData
{
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

   // =DOC SECTION Member eCommodity: COMMODITY
   // =DOC SECTION Enumerated commodity type
   // =DOC TEXT eCommodity: COMMODITY
   // =DOC TEXT Enumerated commodity type
   //////////////////////////////////////////////////////////////////
   COMMODITY eCommodity; 
   
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
   
   // =DOC SECTION Member dTargetConcentrationTime: double
   // =DOC SECTION Target concentration-time (g-hr/m^3)
   // =DOC TEXT dTargetConcentrationTime: double
   // =DOC TEXT Target concentration-time (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dTargetConcentrationTime;

   // =DOC SECTION Member dAchievedConcentrationTime: double
   // =DOC SECTION Calculated achieved concentration-time (g-hr/m^3)
   // =DOC TEXT dAchievedConcentrationTime: double
   // =DOC TEXT Calculated achieved concentration-time (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dAchievedConcentrationTime;

   // =DOC SECTION Member dAchievedDose: double
   // =DOC SECTION Calculated achieved dose (g-hr/m^3)
   // =DOC TEXT dAchievedDose: double
   // =DOC TEXT Calculated achieved dose (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dAchievedDose;

   // =DOC SECTION Member dHlt: double
   // =DOC SECTION Calculated half-loss-time, HLT (hrs)
   // =DOC TEXT dHlt: double
   // =DOC TEXT Calculated half-loss-time, HLT (hrs)
   //////////////////////////////////////////////////////////////////
   double dHlt;
   
   // =DOC SECTION Member dProjectedConcentrationTime: double
   // =DOC SECTION Calculated projected concentration-time (g-hr/m^3)
   // =DOC TEXT dProjectedConcentrationTime: double
   // =DOC TEXT Calculated projected concentration-time (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dProjectedConcentrationTime;

   // =DOC SECTION Member dProjectedTime: double
   // =DOC SECTION Calculated projected for the fumigation (hrs)
   // =DOC TEXT dProjectedTime: double
   // =DOC TEXT Calculated projected for the fumigation (hrs)
   //////////////////////////////////////////////////////////////////
   double dProjectedTime;

   // =DOC SECTION Member dMeanNTC: double
   // =DOC SECTION Calculated projected average target concentration (g/m^3)
   // =DOC TEXT dMeanNTC: double
   // =DOC TEXT Calculated projected average target concentration (g/m^3)
   //////////////////////////////////////////////////////////////////
   double dMeanNTC;

   // =DOC SECTION Member nStatus: int
   // =DOC SECTION Status indicator (0 - Unknown, 1 - Within target, 2 - Above target, 3 - Below target)
   // =DOC TEXT nStatus: int
   // =DOC TEXT Status indicator (0 - Unknown, 1 - Within target, 2 - Above target, 3 - Below target)
   //////////////////////////////////////////////////////////////////
   int nStatus;

   // =DOC SECTION Member dAddProfume: double
   // =DOC SECTION Calculated ProFume to be added to the fumigation (kg)
   // =DOC TEXT dAddProfume: double
   // =DOC TEXT Calculated ProFume to be added to the fumigation (kg)
   //////////////////////////////////////////////////////////////////
   double dAddProfume;

   // =DOC SECTION Member padataAreaMonPt: AREA_MON_PT_DATA_ARRAY*
   // =DOC SECTION Array of area monitor point data
   // =DOC TEXT padataAreaMonPt: AREA_MON_PT_DATA_ARRAY*
   // =DOC TEXT Array of area monitor point data
   //////////////////////////////////////////////////////////////////
   AREA_MON_PT_DATA_ARRAY* padataAreaMonPt;

   // =DOC SECTION Member strMsg: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strMsg: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strMsg;

} AREA_DATA, *LPAREA_DATA;

#endif // !defined(_AREADATA_H_)
