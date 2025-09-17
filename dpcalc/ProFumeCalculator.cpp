// ProFumeCalculator.cpp : Implementation of CProFumeCalculator
#include "stdafx.h"
#include "Dpcalc.h"

#include "Area.h"
#include "DasCommodityData.h"
#include "DasPestData.h"
#include "DoseCalc.h"
#include "GlobalEnums.h"
#include "MonitorPoint.h"
#include "ProFumeCalculator.h"
#include "ValidateAreaMonPtInput.h"
#include "ValidateDoseInput.h"
#include "ValidateHalfLifeInput.h"
#include "ValidateConcentration.h"



/////////////////////////////////////////////////////////////////////////////
// CProFumeCalculator

CProFumeCalculator::CProFumeCalculator()
{
   // Populate the string arrays
   PopulatePest();
   PopulateCommodity();
   PopulateLifeStage();
   PopulateFumigationType();
   PopulateTreatmentType();

   m_padataDose = new DOSE_DATA_ARRAY;

   m_pvalidateDoseInput = new DValidateDoseInput();

   m_pcalcDose = new DDoseCalc();

   m_nMaxProfumeIndex = -1;

   m_pvalidateHalfLifeInput = new DValidateHalfLifeInput();

   m_pcalcMonPt = new DMonitorPoint();

   m_padataConc = new CONCENTRATION_DATA_ARRAY;

   m_pvalidateConc = new DValidateConcentration();

   m_padataAreaMonPt = new AREA_MON_PT_DATA_ARRAY;

   m_pvalidateAreaMonPt = new DValidateAreaMonPtInput();

   m_pcalcArea = new DArea();
}


CProFumeCalculator::~CProFumeCalculator()
{
   CleanPest();
   CleanCommodity();
   CleanFumigationType();
   CleanLifeStage();
   CleanTreatmentType();

   if( m_padataDose )
   {
      m_padataDose->RemoveAll();
   
      delete m_padataDose;
      m_padataDose = NULL;
   }

   if( m_pvalidateDoseInput )
   {
      delete m_pvalidateDoseInput;
      m_pvalidateDoseInput = NULL;
   }

   if( m_pcalcDose )
   {
      delete m_pcalcDose;
      m_pcalcDose = NULL;
   }

   if( m_pvalidateHalfLifeInput )
   {
      delete m_pvalidateHalfLifeInput;
      m_pvalidateHalfLifeInput = NULL;
   }

   if( m_pcalcMonPt )
   {
      delete m_pcalcMonPt;
      m_pcalcMonPt = NULL;
   }

   if( m_padataConc )
   {
      m_padataConc->RemoveAll();

      delete m_padataConc;
      m_padataConc = NULL;
   }

   if( m_pvalidateConc )
   {
      delete m_pvalidateConc;
      m_pvalidateConc = NULL;
   }

   if( m_padataAreaMonPt )
   {
      m_padataAreaMonPt->RemoveAll();

      delete m_padataAreaMonPt;
      m_padataAreaMonPt = NULL;
   }

   if( m_pvalidateAreaMonPt )
   {
      delete m_pvalidateAreaMonPt;
      m_pvalidateAreaMonPt = NULL;
   }

   if( m_pcalcArea )
   {
      delete m_pcalcArea;
      m_pcalcArea = NULL;
   }
}


STDMETHODIMP CProFumeCalculator::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IProFumeCalculator
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
      if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CProFumeCalculator::GetPestTotal( int* Total )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   *Total = m_astrPest.GetSize();

	return S_OK;
}


void CProFumeCalculator::PopulatePest()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   // Get the pest names from the DAS pest data class
   DDasPestData* pDasPestData = NULL;
   pDasPestData = new DDasPestData();

   pDasPestData->GetPestNames( m_astrPest );

   delete pDasPestData;
   pDasPestData = NULL;
}


void CProFumeCalculator::CleanPest()
{
   m_astrPest.RemoveAll();
}



void CProFumeCalculator::PopulateCommodity()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   // Get the commodity names from the commodity data class
   DDasCommodityData* pDasCommodityData = NULL;
   pDasCommodityData = new DDasCommodityData();

   pDasCommodityData->GetCommodityNames( m_astrCommodity );

   delete pDasCommodityData;
   pDasCommodityData = NULL;
}


void CProFumeCalculator::CleanCommodity()
{
   m_astrCommodity.RemoveAll();
}


STDMETHODIMP CProFumeCalculator::GetPestString( int Index, BSTR *Pest )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   CString strPest = m_astrPest.GetAt( Index );

   *Pest = strPest.AllocSysString();

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::GetPestIndex( BSTR *Pest, int *Index )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   HRESULT hr = S_FALSE;

   CString strPest;
   AfxBSTR2CString( &strPest, *Pest );

   int nTotal = m_astrPest.GetSize();

   CString strTempPest;

   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      strTempPest = m_astrPest.GetAt( nIndex );
      
      if( !strPest.CompareNoCase( strTempPest ) )
      {
         *Index = nIndex;
         hr = S_OK;

         break;
      }
   }

	return hr;
}


STDMETHODIMP CProFumeCalculator::SetDosePest( int PestIndex,
                                              int FumigationType,
                                              int LifeStage, 
                                              int TreatmentType,
                                              int Commodity,
                                              double HLT, 
                                              double Temp, 
                                              double TimeExposure, 
                                              double Volume, 
											  double LoadFactor,
											  int IsComplete,
											  int ICounter,
                                              BSTR* Error )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here

   // Validate the input data

   DOSE_INPUT_DATA dataDoseInput;
   
   // Place the data into our validation structure
   dataDoseInput.nPest = PestIndex;
   dataDoseInput.nFumigationType = FumigationType;
   dataDoseInput.nLifeStage = LifeStage;
   dataDoseInput.nTreatmentType = TreatmentType;
   dataDoseInput.nCommodity = Commodity;
   dataDoseInput.dHlt = HLT;
   dataDoseInput.dTemp = Temp;
   dataDoseInput.dTimeExposure = TimeExposure;
   dataDoseInput.dVolume = Volume;
   dataDoseInput.dLoadFactor = LoadFactor;
   dataDoseInput.intIsComplete = IsComplete;
   dataDoseInput.intCounter = ICounter;

   if( !m_pvalidateDoseInput->ValidateDoseInput( &dataDoseInput ) )
   {
      // There was a problem, so return an error
      *Error = dataDoseInput.strError.AllocSysString();

      return S_FALSE;
   }

   // Everything is ok, add the inputs to the dose data array
   DOSE_DATA dataDose;

   dataDose.ePest = static_cast< PEST >( dataDoseInput.nPest );
   dataDose.eFumigationType = static_cast< FUMIGATION_TYPE >( dataDoseInput.nFumigationType );
   dataDose.eLifeStage = static_cast< LIFE_STAGE >( dataDoseInput.nLifeStage );
   dataDose.eTreatmentType = static_cast< TREATMENT_TYPE >( dataDoseInput.nTreatmentType );
   dataDose.eCommodity = static_cast< COMMODITY >( dataDoseInput.nCommodity );
   dataDose.dHltOriginal = dataDoseInput.dHlt;
   dataDose.dTemp = dataDoseInput.dTemp;
   dataDose.dTimeExposure = dataDoseInput.dTimeExposure;
   dataDose.dVolume = dataDoseInput.dVolume;
   dataDose.dProfumeReq = 0.0;
   dataDose.dCylindersReq = 0.0;
   dataDose.dInitialConcentration = 0.0;
   dataDose.dLoadFactor = dataDoseInput.dLoadFactor;
   dataDose.intIsComplete = dataDoseInput.intIsComplete;
   dataDose.intCounter = dataDoseInput.intCounter;

   m_padataDose->Add( dataDose );

   return S_OK;
}

/* 
	This is the subroutine that is called by the profume executable
	It passes the following and retrieves the data back on the same variables it usese. 
	double* ProfumeReq, = 0
	double* Cylinders,  = 0 
	double* Co,			= 0
	double* CT,			= 0
	double* Dose,       = 0
	BSTR*   Error       = ""
	the variables are used as place holders to return data to the Profume application 
*/
  
STDMETHODIMP CProFumeCalculator::GetProfumeRequired( double* ProfumeReq, 
                                                     double* Cylinders, 
                                                     double* Co, 
                                                     double* CT, 
                                                     double* Dose, 
                                                     BSTR*   Error )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here

   // Iterate through the dose data array and calc the profume required
   int nTotal = m_padataDose->GetSize();

   DOSE_DATA dataDose;
   
   double dMaxProfumeReq = -1.0;
   double dMaxDose       = -1.0;
//07/19/07  Daivid Smith added variable for calculation of max exposure time
   double dMaxProfumeCt	 = -1.0;




   m_nMaxProfumeIndex    = -1;
   CString strWarning;

   // Check what the temperature is for the calculation
   // If it's less than 20°C then add a warning message
   dataDose = m_padataDose->GetAt( 0 );

   if( dataDose.dTemp < MIN_PEST_TEMP )
   {
      strWarning.LoadString( IDS_WARNING_MIN_PEST_TEMP );
   }

   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      dataDose = m_padataDose->GetAt( nIndex ); //returns The array element currently at this index

	  //12/23/04-YYT- Assign the passed parameter *CT to the datadose struc
	  //if CT is = 0, it means there is NO user-defined CT, so we compute
	  //else, we use the user-defined CT to calculate.
	  dataDose.dConcentrationTime = *CT;
	
      if( !m_pcalcDose->GetProfumeRequired( &dataDose ) )
      {
         // There was a problem, so return an error right away
         *Error = dataDose.strError.AllocSysString();

         return S_FALSE;
      }

	   // Add any warning messages to the error string
      if( !dataDose.strError.IsEmpty() )
      {
         // Before adding the warning to the existing error message
         // let's ensure the message does not already exist.
         // We don't want to have duplicate error messages
         if( strWarning.Find( dataDose.strError ) < 0 )
         {
            // Ok, the general warning does not exist in our current warning message,
            // However, there may be an embedded warning that should not be duplicated
            // That warning deals with max items being exceeded
            RemoveDuplicateWarning( IDS_WARNING_CO_EXCEEDED, dataDose.strError, strWarning );
            RemoveDuplicateWarning( IDS_WARNING_CT_EXCEEDED, dataDose.strError, strWarning );
            RemoveDuplicateWarning( IDS_WARNING_CO_CT_EXCEEDED, dataDose.strError, strWarning );
            RemoveDuplicateWarning( IDS_WARNING_MIN_PEST_TEMP, dataDose.strError, strWarning );

            if( !strWarning.IsEmpty() )
            {
               strWarning += _T( "\r\n" );
            }

            strWarning += dataDose.strError;
         }
      }

      m_padataDose->SetAt( nIndex, dataDose );

      // Check if this is the worst pest
   /*/  if( dataDose.dProfumeReq >= dMaxProfumeReq )
      {
         if( dataDose.dProfumeReq == dMaxProfumeReq )
         {
            // Since the profume required is the same, lets ensure we select the 
            // worst pest based on the dose required
            // This occurs when the Co is maxed out at 128
            if( dataDose.dDose > dMaxDose )
            {
               dMaxProfumeReq = dataDose.dProfumeReq;
               dMaxDose = dataDose.dDose;

               m_nMaxProfumeIndex = nIndex;
            }
            
            // The dose is not greater than our current value, so don't update our values
         }
         else
         {
            // The amount of profume required is greater that our
            // original value, so let's update the final index value
            dMaxProfumeReq = dataDose.dProfumeReq;
            dMaxDose = dataDose.dDose;

            m_nMaxProfumeIndex = nIndex;
         }
      }*/

      // Checks for max exposure time
	  if(dataDose.dConcentrationTime  > dMaxProfumeCt)
		{
		  {
            // The amount of profume required is greater that our
            // original value, so let's update the final index value
            dMaxProfumeCt = dataDose.dConcentrationTime;
            m_nMaxProfumeIndex = nIndex;
         }
	  
	  
		}







   }// End of For Loop

   // Return the max value back to the caller
   if( m_nMaxProfumeIndex > -1 )
   {
      dataDose = m_padataDose->GetAt( m_nMaxProfumeIndex );

      *ProfumeReq = dataDose.dProfumeReq;
      *Cylinders  = dataDose.dCylindersReq;
      *Co         = dataDose.dInitialConcentration;
      *CT         = dataDose.dConcentrationTime;
      *Dose       = dataDose.dDose;

	  // Jett 6/10/03 Keep C0 in our class m_dInitConc member variable. We will pass this
	  // to Part II later on
	  m_dInitConc = dataDose.dInitialConcentration;
   }
   else
   {
      *ProfumeReq = 0.0;
      *Cylinders = 0.0;
      *Co = 0.0;
      *CT = 0.0;
      *Dose = 0.0;

	  m_dInitConc = 0;
	  // 09/29/09 in case of rodent, m_nMaxProfumeIndex is not being initialized
	  // and causing error in GetMonitorPointStatus at line 533
	  // dataDose = m_padataDose->GetAt( m_nMaxProfumeIndex );
	  m_nMaxProfumeIndex = 0; 
   }

   // Send out the warning messages to the caller
   *Error = strWarning.AllocSysString();
   
   return S_OK;
}


STDMETHODIMP CProFumeCalculator::ClearDosePest()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   m_padataDose->RemoveAll();

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetMonitorPointStatus(double OriginalHLT, 
                                                       double* AchievedCT,
                                                       double* AchievedDose,
                                                       double* ActualOHLT,
                                                       double* ProjectedCT,
                                                       double* ProjectedFumigationTime,
                                                       double* FutureTargetConcentration,
                                                       double* FutureConcentration,
                                                       double* ElapsedTime,
                                                       DATE* LastPeakDateTime,
                                                       DATE* HLTDateTime, 
                                                       BSTR* Msg)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   // Validate the input data (original half life)
   HALFLIFE_INPUT_DATA dataHalfLifeInput;
   dataHalfLifeInput.dHlt = OriginalHLT;

   if( !m_pvalidateHalfLifeInput->ValidateHalfLifeInput( &dataHalfLifeInput ) )
   {
      // There was a problem, so return an error right away
      *Msg = dataHalfLifeInput.strMsg.AllocSysString();

      return S_FALSE;
   }

   MONITOR_POINT_DATA dataMonPt;
   dataMonPt.dHltOriginal = dataHalfLifeInput.dHlt;
   dataMonPt.padataConc = m_padataConc;
   
   // We know what the worst case pest is so we will use it's data
   DOSE_DATA dataDose;
   dataDose = m_padataDose->GetAt( m_nMaxProfumeIndex );

   dataMonPt.ePest = dataDose.ePest;
   dataMonPt.eFumigationType = dataDose.eFumigationType;
   dataMonPt.eLifeStage = dataDose.eLifeStage;
   dataMonPt.eTreatmentType = dataDose.eTreatmentType;
   dataMonPt.dTemp = dataDose.dTemp;
   dataMonPt.dVolume = dataDose.dVolume;
   dataMonPt.dTimeExposure = dataDose.dTimeExposure;
   
   // Get the half-life for the current concentration
   // Jett 6/10/03 passing the C0 values we got from Part I Step 7
   if(m_pcalcMonPt->GetMonitorPointStatus(&dataMonPt, m_dInitConc, OriginalHLT))
   {
      *AchievedDose              = dataMonPt.dAchievedDose;
      *AchievedCT                = dataMonPt.dAchievedCT;
      *ActualOHLT                = dataMonPt.dActualOHLT;
      *ProjectedCT               = dataMonPt.dProjectedCT;
      *ProjectedFumigationTime   = dataMonPt.dProjectedFumigationTime;
      *FutureTargetConcentration = dataMonPt.dNewTargetC;
      *FutureConcentration       = dataMonPt.dFutureC;
      *ElapsedTime               = dataMonPt.odtsLastCTime.GetTotalSeconds() / 3600;
      *LastPeakDateTime          = dataMonPt.odtLastPeakDateTime;
      *HLTDateTime               = dataMonPt.odtHltDateTime;
   }
   else
   {
      *AchievedDose              = -1.0;
      *AchievedCT                = -1.0;
      *ActualOHLT                = -1.0;
      *ProjectedCT               = -1.0;
      *ProjectedFumigationTime   = -1.0;
      *FutureTargetConcentration = -1.0;
      *FutureConcentration       = -1.0;
      *ElapsedTime               = -1.0;
      *LastPeakDateTime          = 0.0;
      *HLTDateTime               = 0.0;

      return S_FALSE;
   }

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::SetConcentration( double Concentration, DATE Date, BSTR* Msg )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   
   // Validate the values
   CONC_INPUT_DATA dataConcInput;

   dataConcInput.dConcentration = Concentration;
   dataConcInput.odtDateTime = Date;

   if( !m_pvalidateConc->ValidateConcentration( &dataConcInput ) )
   {
      *Msg = dataConcInput.strMsg.AllocSysString();

      return S_FALSE;
   }

   // Add the values to the array
   CONCENTRATION_DATA dataConc;
   dataConc.dConcentration = dataConcInput.dConcentration;
   dataConc.odtDateTime = dataConcInput.odtDateTime;

   m_padataConc->Add( dataConc );

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::ClearConcentrations()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   m_padataConc->RemoveAll();

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::GetVersion(BSTR *Version)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   //JNM 05/13/03 - Changed version from 1.02
   //TSG 05/16/03 - Changed version from 2.01
   //TSG 06/10/03 - 2.04 - change in calculation for CnTTargetDosage
   //CHG 09/02/03 - 2.0.5 - change in the introduction rate formula for the 6.35mm
   //CHG 09/02/03 - 2.0.6 - change in the introduction rate formula for the 6.35mm
   //CHG 01/28/04 - 2.0.7 - Launch Changes (life stage wording, safety factor, 
   //                       space/vacuum max ct)
   //CHG 01/28/04 - 2.1 - Launch Changes
   //CHG 03/03/04 - 2.2 - Ticket # 2120829 - Safety Factor Changes
   //CHG 09/29/09 - 2.2.526 - Error out while using Rodent
   CString strVersion = _T( "2.2.526" );
   *Version = strVersion.AllocSysString();

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::GetPt2PtHlt( double* LastPeakHLT,
                                              DATE* LastPeakDateTime,
                                              DATE* HLTDateTime, 
                                              BSTR* Msg )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   MONITOR_POINT_2_POINT_DATA dataMonPt2Pt;

   // Pass the concentration data
   dataMonPt2Pt.padataConc = m_padataConc;

   // Get the half-life for the current concentration
   if( m_pcalcMonPt->GetPt2PtHlt( &dataMonPt2Pt ) )
   {
      *LastPeakHLT = dataMonPt2Pt.dActualOHLT;
      *LastPeakDateTime = dataMonPt2Pt.odtLastPeakDateTime;
      *HLTDateTime = dataMonPt2Pt.odtLastPeakDateTime;
   }
   else
   {
      *LastPeakHLT = -1.0;
      *LastPeakDateTime = 0.0;
      *HLTDateTime = 0.0;

      *Msg = dataMonPt2Pt.strMsg.AllocSysString();

      return S_FALSE;
   }

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::SetAreaMonitorPoint( double AchievedConcentrationTime, 
                                                      double AchievedDose,
                                                      double ElapsedTime,
                                                      double FutureTargetConcentration,
                                                      double FutureConcentration,
                                                      double Hlt,
                                                      double ProjectedConcentrationTime,
                                                      double ProjectedTime,
                                                      BSTR* Msg )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   AREA_MON_PT_DATA dataAreaMonPt;
   dataAreaMonPt.dAchievedConcentrationTime = AchievedConcentrationTime;
   dataAreaMonPt.dAchievedDose = AchievedDose;
   dataAreaMonPt.dElapsedTime = ElapsedTime;
   dataAreaMonPt.dFutureTargetConcentration = FutureTargetConcentration;
   dataAreaMonPt.dFutureConcentration = FutureConcentration;
   dataAreaMonPt.dHlt = Hlt;
   dataAreaMonPt.dProjectedConcentrationTime = ProjectedConcentrationTime;
   dataAreaMonPt.dProjectedTime = ProjectedTime;

   // Check the input data
   if( !m_pvalidateAreaMonPt->ValidateAreaMonPt( &dataAreaMonPt ) )
   {
      *Msg = dataAreaMonPt.strMsg.AllocSysString();

      return S_FALSE;
   }

   // Add the data to the array

   m_padataAreaMonPt->Add( dataAreaMonPt );

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::ClearAreaMonitorPoints()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   m_padataAreaMonPt->RemoveAll();

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::GetAreaStatus( double  TargetConcentrationTime,
                                                double* AchievedConcentrationTime, 
                                                double* AchievedDose,
                                                double* Hlt,
                                                double* ProjectedConcentrationTime,
                                                double* ProjectedTime,
                                                double* MeanNTC,
                                                int*    Status,
                                                double* AddProfume,
                                                BSTR*   Msg )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   AREA_DATA dataArea;
   dataArea.dTargetConcentrationTime = TargetConcentrationTime;
   dataArea.padataAreaMonPt = m_padataAreaMonPt;

   // We know what the worst case pest is, so we will use its data
   DOSE_DATA dataDose;
   
   if( m_nMaxProfumeIndex >= 0 )
   {
      dataDose = m_padataDose->GetAt( m_nMaxProfumeIndex );
   }
   else
   {
      *AchievedConcentrationTime = -1.0;
      *AchievedDose = -1.0;
      *Hlt = -1;
      *ProjectedConcentrationTime = -1.0;
      *ProjectedTime = -1.0;
      *MeanNTC = -1.0;
      *Status = STATUS_UNKNOWN;
      *AddProfume = -1.0;
      
      return S_OK;
   }

   // 5/16/2003 Jett - this is the only exception wherein the status if forced upon us
   // by the frontend. Frontend is responsible for validating if we have readings for
   // at all monitoring points at all times. If it finds out that we do not have a
   // complete set of readings, it passes a STATUS_MISSING_CONC_READING and we return
   // the same status and all return values set to negative one.
   // (We also need to check if any monitoring points were passed, if no area monitor
   // points were passed then we definitely did not have a reading)
   if (*Status == STATUS_MISSING_CONC_READING && m_padataAreaMonPt->GetSize() == 0)
   {
      *AchievedConcentrationTime = -1.0;
      *AchievedDose = -1.0;
      *Hlt = -1;
      *ProjectedConcentrationTime = -1.0;
      *ProjectedTime = -1.0;
      *MeanNTC = -1.0;
      *Status = STATUS_MISSING_CONC_READING;
      *AddProfume = -1.0;
      
      return S_OK;
   }

   dataArea.ePest = dataDose.ePest;
   dataArea.eLifeStage = dataDose.eLifeStage;
   dataArea.eTreatmentType = dataDose.eTreatmentType;
   dataArea.eCommodity = dataDose.eCommodity;
   dataArea.dTemp = dataDose.dTemp;
   dataArea.dTimeExposure = dataDose.dTimeExposure;
   dataArea.dVolume = dataDose.dVolume;

   if( m_pcalcArea->GetAreaStatus( &dataArea ) )
   {
      *AchievedConcentrationTime = dataArea.dAchievedConcentrationTime;
      *AchievedDose = dataArea.dAchievedDose;
      *Hlt = dataArea.dHlt;
      *ProjectedConcentrationTime = dataArea.dProjectedConcentrationTime;
      *ProjectedTime = dataArea.dProjectedTime;
      *MeanNTC = dataArea.dMeanNTC;
      *Status = dataArea.nStatus;
      *AddProfume = dataArea.dAddProfume;
   }
   else
   {
      *AchievedConcentrationTime = -1.0;
      *AchievedDose = -1.0;
      *Hlt = -1;
      *ProjectedConcentrationTime = -1.0;
      *ProjectedTime = -1.0;
      *MeanNTC = -1.0;
      *Status = STATUS_UNKNOWN;
      *AddProfume = -1.0;

      *Msg = dataArea.strMsg.AllocSysString();
      
      return S_FALSE;
   }

	return S_OK;
}

void CProFumeCalculator::RemoveDuplicateWarning( UINT unWarningId, 
                                                 CString &strNewWarning, 
                                                 CString &strExistingWarning )
{
   CString strMaxExceeded;
   strMaxExceeded.LoadString( unWarningId );
   
   if( strNewWarning.Find( strMaxExceeded ) >= 0 && strExistingWarning.Find( strMaxExceeded ) >= 0 )
   {
      // Need to remove the warning from the dataDose.strError
      strNewWarning.Replace( strMaxExceeded, _T( "" ) );
   }

   // Before adding the error, check if it ends in a \r\n and remove it
   if( strNewWarning.Find( _T( "\r\n" ) ) >= 0 )
   {
      strNewWarning.Replace( _T( "\r\n" ), _T( "" ) );
   }
}


STDMETHODIMP CProFumeCalculator::GetCommodityIndex( BSTR *Commodity, int *Index )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   HRESULT hr = S_FALSE;

   CString strCommodity;
   AfxBSTR2CString( &strCommodity, *Commodity );

   int nTotal = m_astrCommodity.GetSize();

   CString strTempCommodity;

   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      strTempCommodity = m_astrCommodity.GetAt( nIndex );
      
      if( !strCommodity.CompareNoCase( strTempCommodity ) )
      {
         *Index = nIndex;
         hr = S_OK;

         break;
      }
   }

	return hr;
}


STDMETHODIMP CProFumeCalculator::GetCommodityString( int Index, BSTR *Commodity )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   CString strCommodity = m_astrCommodity.GetAt( Index );

   *Commodity = strCommodity.AllocSysString();

	return S_OK;
}


STDMETHODIMP CProFumeCalculator::GetCommodityTotal( int *Total )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   *Total = m_astrCommodity.GetSize();

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetLifeStageTotal( int *Total )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   *Total = m_astrLifeStage.GetSize();

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetLifeStageIndex( BSTR *LifeStage, int *Index )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   HRESULT hr = S_FALSE;

   CString strLifeStage;
   AfxBSTR2CString( &strLifeStage, *LifeStage );

   int nTotal = m_astrLifeStage.GetSize();

   CString strTempLifeStage;

   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      strTempLifeStage = m_astrLifeStage.GetAt( nIndex );
      
      if( !strLifeStage.CompareNoCase( strTempLifeStage ) )
      {
         *Index = nIndex;
         hr = S_OK;

         break;
      }
   }

	return hr;

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetLifeStageString( int Index, BSTR *LifeStage )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   CString strLifeStage = m_astrLifeStage.GetAt( Index );

   *LifeStage = strLifeStage.AllocSysString();

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetFumigationTypeTotal( int *Total )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   *Total = m_astrFumigationType.GetSize();

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetFumigationTypeIndex( BSTR *FumigationType, int *Index )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   HRESULT hr = S_FALSE;

   CString strFumigationType;
   AfxBSTR2CString( &strFumigationType, *FumigationType );

   int nTotal = m_astrFumigationType.GetSize();

   CString strTempFumigationType;

   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      strTempFumigationType = m_astrFumigationType.GetAt( nIndex );
      
      if( !strFumigationType.CompareNoCase( strTempFumigationType ) )
      {
         *Index = nIndex;
         hr = S_OK;

         break;
      }
   }

	return hr;

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetFumigationTypeString( int Index, BSTR *FumigationType )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   CString strFumigationType = m_astrFumigationType.GetAt( Index );

   *FumigationType = strFumigationType.AllocSysString();

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetTreatmentTypeTotal( int *Total )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   *Total = m_astrTreatmentType.GetSize();

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetTreatmentTypeIndex( BSTR *TreatmentType, int *Index )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   HRESULT hr = S_FALSE;

   CString strTreatmentType;
   AfxBSTR2CString( &strTreatmentType, *TreatmentType );

   int nTotal = m_astrTreatmentType.GetSize();

   CString strTempTreatmentType;

   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      strTempTreatmentType = m_astrTreatmentType.GetAt( nIndex );
      
      if( !strTreatmentType.CompareNoCase( strTempTreatmentType ) )
      {
         *Index = nIndex;
         hr = S_OK;

         break;
      }
   }

	return hr;

	return S_OK;
}

STDMETHODIMP CProFumeCalculator::GetTreatmentTypeString( int Index, BSTR *TreatmentType )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
   CString strTreatmentType = m_astrTreatmentType.GetAt( Index );

   *TreatmentType = strTreatmentType.AllocSysString();

	return S_OK;
}

void CProFumeCalculator::PopulateFumigationType()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   // Get the commodity names from the commodity data class
   DDasFumigationTypeData* pDasFumigationTypeData = NULL;
   pDasFumigationTypeData = new DDasFumigationTypeData();

   pDasFumigationTypeData->GetFumigationTypeNames( m_astrFumigationType );

   delete pDasFumigationTypeData;
   pDasFumigationTypeData = NULL;
}

void CProFumeCalculator::CleanFumigationType()
{
   m_astrFumigationType.RemoveAll();
}

void CProFumeCalculator::PopulateLifeStage()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   // Get the commodity names from the commodity data class
   DDasLifeStageData* pDasLifeStageData = NULL;
   pDasLifeStageData = new DDasLifeStageData();

   pDasLifeStageData->GetLifeStageNames( m_astrLifeStage );

   delete pDasLifeStageData;
   pDasLifeStageData = NULL;
}

void CProFumeCalculator::CleanLifeStage()
{
   m_astrLifeStage.RemoveAll();
}

void CProFumeCalculator::PopulateTreatmentType()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

   // Get the commodity names from the commodity data class
   DDasTreatmentTypeData* pDasTreatmentTypeData = NULL;
   pDasTreatmentTypeData = new DDasTreatmentTypeData();

   pDasTreatmentTypeData->GetTreatmentTypeNames( m_astrTreatmentType );

   delete pDasTreatmentTypeData;
   pDasTreatmentTypeData = NULL;
}

void CProFumeCalculator::CleanTreatmentType()
{
   m_astrTreatmentType.RemoveAll();
}
