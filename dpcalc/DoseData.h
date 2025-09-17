// =DOC FILE Structure for dose information
// =DOC TEXT Declares tagDoseData structure and an array of tagDoseData structures
//////////////////////////////////////////////////////////////////

// DoseData.h
//
//////////////////////////////////////////////////////////////////////

#if !defined(_DOSEDATA_H_)
#define _DOSEDATA_H_


#include <afxtempl.h>
#include "DasCommodityData.h"
#include "DasFumigationTypeData.h"
#include "DasLifeStageData.h"
#include "DasPestData.h"
#include "DasTreatmentTypeData.h"



// =DOC SECTION Struct tagDoseData ( DOSE_DATA, *LPDOSE_DATA )
// =DOC SECTION Carrier structure containing dose information, both input and calculated
// =DOC TEXT struct tagDoseData extends nothing
// =DOC TEXT Carrier structure containing dose information, both input and calculated
//////////////////////////////////////////////////////////////////
typedef struct tagDoseData
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

   // =DOC SECTION Member dHltOriginal: double
   // =DOC SECTION User entered HLT (hrs)
   // =DOC TEXT dHltOriginal: double
   // =DOC TEXT User entered HLT (hrs)
   //////////////////////////////////////////////////////////////////
   double dHltOriginal;

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

   // =DOC SECTION Member dInitialConcentration: double
   // =DOC SECTION Calculated maximum initial concentration (g/m^3)
   // =DOC TEXT dInitialConcentration: double
   // =DOC TEXT Calculated maximum initial concentration (g/m^3)
   //////////////////////////////////////////////////////////////////
   double dInitialConcentration;
   double dInitialConcentration_FOR_CT;
   // =DOC SECTION Member dProfumeReq: double
   // =DOC SECTION Calculated ProFume required (kg)
   // =DOC TEXT dProfumeReq: double
   // =DOC TEXT Calculated ProFume required (kg)
   //////////////////////////////////////////////////////////////////
   double dProfumeReq;
   
   // =DOC SECTION Member dCylindersReq: double
   // =DOC SECTION Calculated number of ProFume cylinders required
   // =DOC TEXT dCylindersReq: double
   // =DOC TEXT Calculated number of ProFume cylinders required
   //////////////////////////////////////////////////////////////////
   double dCylindersReq;

   // =DOC SECTION Member dConcentrationTime: double
   // =DOC SECTION Calculated concentration time, CT (g-hr/m^3)
   // =DOC TEXT dConcentrationTime: double
   // =DOC TEXT Calculated concentration time, CT (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dConcentrationTime;

   // =DOC SECTION Member dDose: double
   // =DOC SECTION Calculated target dose (g-hr/m^3)
   // =DOC TEXT dDose: double
   // =DOC TEXT Calculated target dose (g-hr/m^3)
   //////////////////////////////////////////////////////////////////
   double dDose;
   
   // =DOC SECTION Member dLoadFactor: double
   // =DOC SECTION Inputted data from the Dosage Plan tab
   // =DOC TEXT dLoadFactor: double
   // =DOC TEXT Inputted data from the Dosage Plan tab
   //////////////////////////////////////////////////////////////////
   double dLoadFactor;
   
   // =DOC SECTION Member intIsComplete: integer
   // =DOC SECTION This would indicate if the criteria are complete to compute for HLT sorption
   // =DOC TEXT intIsComplete: integer
   // =DOC TEXT This would indicate if the criteria are complete to compute for HLT sorption
   //////////////////////////////////////////////////////////////////
   int intIsComplete;

   // =DOC SECTION Member nCounter: integer
   // =DOC SECTION This would indicate if we are now in the next record
   // =DOC TEXT nCounter: integer
   // =DOC TEXT This would indicate if we are now in the next record
   //////////////////////////////////////////////////////////////////
   int intCounter;
   
   // =DOC SECTION Member strError: CString
   // =DOC SECTION The message string if any alerts occur
   // =DOC TEXT strError: CString
   // =DOC TEXT The message string if any alerts occur
   //////////////////////////////////////////////////////////////////
   CString strError;

} DOSE_DATA, *LPDOSE_DATA;


// =DOC SECTION CArray DOSE_DATA_ARRAY
// =DOC SECTION Array of DOSE_DATA structures
// =DOC TEXT CArray DOSE_DATA_ARRAY extends CArray
// =DOC TEXT Array of DOSE_DATA structures
//////////////////////////////////////////////////////////////////
typedef CArray < DOSE_DATA, DOSE_DATA& > DOSE_DATA_ARRAY;

#endif // !defined(_DOSEDATA_H_)
