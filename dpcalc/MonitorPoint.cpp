// MonitorPoint.cpp: implementation of the DMonitorPoint class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "GlobalFunctions.h"
#include "GlobalEnums.h"
#include "MonitorPoint.h"
#include <math.h>



#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif



//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DMonitorPoint::DMonitorPoint()
{

}


DMonitorPoint::~DMonitorPoint()
{

}


// Jett 6/10/03 - modified function declaration so we can pass the Intial concentration and OHLT values from Part I
// this is needed to calculate for the CnTTargetD used to get the DeltaDose
BOOL DMonitorPoint::GetMonitorPointStatus( const LPMONITOR_POINT_DATA lpdataMonPt, const double dInitConc, const double dOHLT )
{
   ASSERT(lpdataMonPt);

   BOOL bSuccess = FALSE;

   CalcInfo highCalcInfo;
   double dHighTemp;
   double dLowTemp;

   //JNM 04/14/03 - Fum. Enh. 2003 - Added new variables below.
   // Concentration Readings for this monitor point
   CONCENTRATION_DATA_ARRAY* pConcentrationArray = lpdataMonPt->padataConc;

   // number of concentration readings
   int nConcentrationReadingCount = pConcentrationArray->GetSize();

   // last concentration reading taken for this monitor point
   double dCurrentConcentration = pConcentrationArray->GetAt( nConcentrationReadingCount - 1 ).dConcentration;

   // date and time for last concentration reading
   COleDateTime odtCurrentTime = pConcentrationArray->GetAt( nConcentrationReadingCount - 1 ).odtDateTime;

   // time at start of exposure
   COleDateTimeSpan odtsExposureTime = odtCurrentTime - pConcentrationArray->GetAt( 0 ).odtDateTime;

   // store our interpolated Delta D
   double dIntDeltaD; 

   //JNM 04/16/03 - Fum. Enh. 2003 - Store the interpolated n value
   double dIntPestNValue;


   // Populate the pest array with the correct pests information.
   m_DasPestData.PopulatePestData( lpdataMonPt->ePest, 
                                   lpdataMonPt->eLifeStage, 
                                   lpdataMonPt->eTreatmentType, 
                                   lpdataMonPt->eFumigationType,
                                   lpdataMonPt->dTimeExposure,
                                   lpdataMonPt->strMsg );

   // Jett 6/5/03 set the C0 value for the high temperature
   highCalcInfo.dInitConc = dInitConc;
   highCalcInfo.dPlannedHLT = dOHLT;

   // Get the current pest data for the high temperature, and perform the calculations.
   this->m_pestCurrent = this->m_DasPestData.GetPestData(lpdataMonPt->dTemp, FALSE);
   dHighTemp = this->m_pestCurrent.dTemp;
   bSuccess = this->DoMonitorPointCalculations(lpdataMonPt, &highCalcInfo);

   if (bSuccess)
   {
      // Get the current pest data for the low temperature, but only perform the calculations
      // if the high temperature and the low temperature are unequal.
      this->m_pestCurrent = this->m_DasPestData.GetPestData(lpdataMonPt->dTemp, TRUE);
      dLowTemp = this->m_pestCurrent.dTemp;

      if( dHighTemp != dLowTemp )
      {
         CalcInfo lowCalcInfo;

		 // Jett 6/5/03 set the C0 value for the low temperature. This is actually
		 // the same as the C0 for the low temp because we interpolated earlier
		 // during Part I
  	     lowCalcInfo.dInitConc = dInitConc;
		 lowCalcInfo.dPlannedHLT = dOHLT;

         bSuccess = this->DoMonitorPointCalculations(lpdataMonPt, &lowCalcInfo);

         // Interpolate the results.
         double dIntTemp = (lpdataMonPt->dTemp - dLowTemp) / (dHighTemp - dLowTemp);
         
		 //JNM 04/14/03 - Fum. Enh. 2003 - Removed the old interpolation method below
		 //		in getting the Projected Time value.
         //if( highCalcInfo.dProjectedFumigationTime < 0.0 || lowCalcInfo.dProjectedFumigationTime < 0.0 )
         //{
         //   lpdataMonPt->dProjectedFumigationTime = -1;
         //}
         //else
         //{
         //   lpdataMonPt->dProjectedFumigationTime = (highCalcInfo.dProjectedFumigationTime - lowCalcInfo.dProjectedFumigationTime) * dIntTemp + lowCalcInfo.dProjectedFumigationTime;
         //}

		 //JNM 04/14/03 - Fum. Enh. 2003 - interpolate DeltaD
		 dIntDeltaD = (highCalcInfo.dDeltaDose - lowCalcInfo.dDeltaDose) * dIntTemp + lowCalcInfo.dDeltaDose;

		 //JNM 04/16/03 - Fum. Enh. 2003 - interpolate the n value which will be used for the
		 //		Extra Time calculation
		 dIntPestNValue = (highCalcInfo.dPestNValue - lowCalcInfo.dPestNValue) * dIntTemp + lowCalcInfo.dPestNValue;

		 lpdataMonPt->dRemainingDose = (highCalcInfo.dRemainingDose - lowCalcInfo.dRemainingDose) * dIntTemp + lowCalcInfo.dRemainingDose;
         lpdataMonPt->dAchievedCT    = (highCalcInfo.dAchievedCT - lowCalcInfo.dAchievedCT) * dIntTemp + lowCalcInfo.dAchievedCT;
         lpdataMonPt->dAchievedDose  = (highCalcInfo.dAchievedDose - lowCalcInfo.dAchievedDose) * dIntTemp + lowCalcInfo.dAchievedDose;
         lpdataMonPt->dActualEHLT    = (highCalcInfo.dActualEHLT - lowCalcInfo.dActualEHLT) * dIntTemp + lowCalcInfo.dActualEHLT;
         lpdataMonPt->dActualOHLT    = (highCalcInfo.dActualOHLT - lowCalcInfo.dActualOHLT) * dIntTemp + lowCalcInfo.dActualOHLT;
         lpdataMonPt->dFutureC       = (highCalcInfo.dFutureC - lowCalcInfo.dFutureC) * dIntTemp + lowCalcInfo.dFutureC;
         lpdataMonPt->dNewTargetC    = (highCalcInfo.dNewTargetC - lowCalcInfo.dNewTargetC) * dIntTemp + lowCalcInfo.dNewTargetC;
         lpdataMonPt->dFutureDose    = (highCalcInfo.dFutureDose - lowCalcInfo.dFutureDose) * dIntTemp + lowCalcInfo.dFutureDose;
         lpdataMonPt->dProjectedCT   = (highCalcInfo.dProjectedCT - lowCalcInfo.dProjectedCT) * dIntTemp + lowCalcInfo.dProjectedCT;
         lpdataMonPt->dRemainingCT   = (highCalcInfo.dRemainingCT - lowCalcInfo.dRemainingCT) * dIntTemp + lowCalcInfo.dRemainingCT;

		  //JNM 04/14/03 - Fum. Enh. 2003 - call the function to calculate Projected time
		 //		which would depend on our concentrations
		 if(highCalcInfo.dActualEHLT == 0.0)
			 //We are in a plateau
			 lpdataMonPt->dProjectedFumigationTime = GetProjectedFumigationTime2(dCurrentConcentration, dIntDeltaD, dIntPestNValue, odtsExposureTime);
		 else if(highCalcInfo.dActualEHLT == -1.0)
			 //We are increasing concentration
			 lpdataMonPt->dProjectedFumigationTime = -1.0;
		 else
			//We are decreasing concentration
			lpdataMonPt->dProjectedFumigationTime = GetProjectedFumigationTime( lpdataMonPt->dActualEHLT, dCurrentConcentration,dIntDeltaD, dIntPestNValue, odtsExposureTime );
		 
      }
      else
      {
         // The high and low temperatures are the same, so just use the high values.
         lpdataMonPt->dProjectedFumigationTime = highCalcInfo.dProjectedFumigationTime;
         lpdataMonPt->dRemainingDose = highCalcInfo.dRemainingDose;
         lpdataMonPt->dAchievedCT   = highCalcInfo.dAchievedCT;
         lpdataMonPt->dAchievedDose = highCalcInfo.dAchievedDose;
         lpdataMonPt->dActualEHLT   = highCalcInfo.dActualEHLT;
         lpdataMonPt->dActualOHLT   = highCalcInfo.dActualOHLT;
         lpdataMonPt->dFutureC      = highCalcInfo.dFutureC;
         lpdataMonPt->dNewTargetC   = highCalcInfo.dNewTargetC;
         lpdataMonPt->dFutureDose   = highCalcInfo.dFutureDose;
         lpdataMonPt->dProjectedCT  = highCalcInfo.dProjectedCT;
         lpdataMonPt->dRemainingCT  = highCalcInfo.dRemainingCT;

         bSuccess = TRUE;
      }

      // Get the time T for the last monitor point reading.
      if( lpdataMonPt->dAchievedCT >= 0.0 )
      {
         CONCENTRATION_DATA_ARRAY* pConcData = lpdataMonPt->padataConc;
         COleDateTime odtStart = pConcData->GetAt(0).odtDateTime;
         COleDateTime odtEnd = pConcData->GetAt(pConcData->GetUpperBound()).odtDateTime;
         lpdataMonPt->odtsLastCTime = odtEnd - odtStart;

         // Get the last peak time.
         this->GetLastPeakIndex(pConcData, lpdataMonPt->odtLastPeakDateTime);

         // Get the last point time
         lpdataMonPt->odtHltDateTime = odtEnd;
      }
      else
      {
         lpdataMonPt->odtsLastCTime = -1;
      }
   }

   return bSuccess;
}


enum PEAK_PATTERN
{
   PP_INCREASING = 1,
   PP_PLATEAUED  = 2,
   PP_DECREASING = 3
};


BOOL DMonitorPoint::DoMonitorPointCalculations(const LPMONITOR_POINT_DATA pMonitorPointData, DMonitorPoint::CalcInfo* pCalcInfo)
{
   ASSERT(pMonitorPointData && pCalcInfo);

   BOOL bRet = FALSE;
   CONCENTRATION_DATA_ARRAY* pConcentrationArray = pMonitorPointData->padataConc;
   int nConcentrationReadingCount = pConcentrationArray->GetSize();

   if (nConcentrationReadingCount > 0)
   {
      // Get our current time and concentration information, since all equations 
      // are based on the current time/concentration
      COleDateTime odtCurrentTime = pConcentrationArray->GetAt( nConcentrationReadingCount - 1 ).odtDateTime;
      double dCurrentConcentration = pConcentrationArray->GetAt( nConcentrationReadingCount - 1 ).dConcentration;

	  //JNM 04/16/03 - Fum. Enh. 2003 - Store the current pest n value in the structure.
	  //	This will be used later in interpolating n for the Extra Time calculation.
	  pCalcInfo->dPestNValue = m_pestCurrent.dNValue;

      // Part II, Step 1
      pCalcInfo->dAchievedCT = GetCTAchieved(pConcentrationArray);
      pCalcInfo->dAchievedDose = GetDoseAchieved(pConcentrationArray);

      // Part II, Step 2
      COleDateTime odtLastMaxTime;
      int nLastMaxIndex = GetLastPeakIndex(pConcentrationArray, odtLastMaxTime);
      
      // Get the Peak Pattern
      PEAK_PATTERN peakPattern = GetPeakPattern(pMonitorPointData->padataConc, nLastMaxIndex);

      if ( peakPattern == PP_DECREASING)
      {
         // Get the actual HLT
         pCalcInfo->dActualOHLT = GetActualOHLT(pMonitorPointData->padataConc, nLastMaxIndex, odtCurrentTime);

         // Use the OHLT to continue with the rest of the steps
         pCalcInfo->dActualEHLT = ::GetEHLT(pCalcInfo->dActualOHLT, m_pestCurrent.dNValue);


         // Part II, Step 3
         COleDateTimeSpan odtsExposureTime = odtCurrentTime - pConcentrationArray->GetAt( 0 ).odtDateTime;
         COleDateTimeSpan odtsProjectedExposureTime( 0, 0, 0, static_cast< int >( pMonitorPointData->dTimeExposure * 3600 ) );
         pCalcInfo->dRemainingCT = GetRemainingCT( dCurrentConcentration, 
                                                   odtsExposureTime, 
                                                   odtsProjectedExposureTime, 
                                                   pCalcInfo->dActualOHLT );

         // Part II, Step 4
         pCalcInfo->dProjectedCT = GetProjectedCT( pCalcInfo->dAchievedCT, pCalcInfo->dRemainingCT );

         // Part II, Steps 5-6
         // This space intentionally left blank.  The calcs that are occurring here are area calculations.

         // Part II, Step 7 & 8
         // Step 7 is for more time and step 8 is for less time
         // The same equations are being used for both

		 // Jett 6/10/03 - change the way we get TargetDose (CnTTargetD). In the past
		 // we have been using TargetD based on Part I Step 4. We should be using the
		 // one in Part I Step 10C

         //double dTargetDose = ::GetFinalDose( m_pestCurrent.dIntercept, 
         //                                     m_pestCurrent.dSlope, 
         //                                     m_pestCurrent.dProbit,
         //                                     m_pestCurrent.dSafetyFactor );

		 // we should really be using the planned exposure time, but since 
		 // plan is always = actual then we can just use projected exposure time
		 double dExposureTime = odtsProjectedExposureTime.GetTotalSeconds() / 3600;

		 double dTargetDose = GetTargetDose (pCalcInfo->dInitConc,
			                                 pCalcInfo->dPestNValue,
											 pCalcInfo->dPlannedHLT,
											 dExposureTime);

         pCalcInfo->dDeltaDose = GetDeltaDose( dTargetDose, pCalcInfo->dAchievedDose );

         COleDateTimeSpan odtsCurrentTime = odtCurrentTime - pConcentrationArray->GetAt( 0 ).odtDateTime;
         CONCENTRATION_DATA lastMaxConcentration = pConcentrationArray->GetAt( nLastMaxIndex );

		 //JNM 04/16/03 - Fum. Enh. 2003 - Pass m_pestCurrent.dNValue to function
         pCalcInfo->dProjectedFumigationTime = GetProjectedFumigationTime( pCalcInfo->dActualEHLT, 
                                                                           dCurrentConcentration, 
                                                                           pCalcInfo->dDeltaDose,
																		   m_pestCurrent.dNValue, 
                                                                           odtsExposureTime );

         // Part II, Step 9
         // When determining if more product needs to be added, we will assume a 1 hour lag time
         // from the most recent monitoring time

         // Part II, Step 9B
         pCalcInfo->dFutureC = GetFutureConcentration( dCurrentConcentration, pCalcInfo->dActualOHLT );

         // Part II, Step 9C
         pCalcInfo->dFutureDose = GetFutureDose( pCalcInfo->dFutureC, dCurrentConcentration );

         // Part II, Step 9D
         pCalcInfo->dRemainingDose = GetRemainingRequiredDose( dTargetDose, 
                                                               pCalcInfo->dAchievedDose, 
                                                               pCalcInfo->dFutureDose );

         // Part II, Step 9E
         pCalcInfo->dNewTargetC = GetNewTargetConcentration( pCalcInfo->dRemainingDose,
                                                             pCalcInfo->dActualEHLT,
                                                             odtsExposureTime, 
                                                             odtsProjectedExposureTime );
      }
      else if ( peakPattern == PP_PLATEAUED )
      {
         // We cannot calculate the actual OHLT or EHLT
         pCalcInfo->dActualOHLT = 0.0;
         pCalcInfo->dActualEHLT = 0.0;

         // Part II, Step 3
         // Use the alternate calculation for the case PP_PLATEAUED
         COleDateTimeSpan odtsExposureTime = odtCurrentTime - pConcentrationArray->GetAt( 0 ).odtDateTime;
         COleDateTimeSpan odtsProjectedExposureTime( 0, 0, 0, static_cast< int >( pMonitorPointData->dTimeExposure * 3600 ) );
         pCalcInfo->dRemainingCT = GetRemainingCT2( dCurrentConcentration, 
                                                    odtsExposureTime, 
                                                    odtsProjectedExposureTime );

         // Part II, Step 4
         pCalcInfo->dProjectedCT = GetProjectedCT( pCalcInfo->dAchievedCT, pCalcInfo->dRemainingCT );

         // Part II, Steps 5-6
         // This space intentionally left blank.  The calcs that are occurring here are area calculations.

         // Part II, Step 7 & 8
         // Step 7 is for more time and step 8 is for less time
         // The same equations are being used for both


		 // Jett 6/10/03 - change the way we get TargetDose (CnTTargetD). In the past
		 // we have been using TargetD based on Part I Step 4. We should be using the
		 // one in Part I Step 10C

         //double dTargetDose = ::GetFinalDose( m_pestCurrent.dIntercept, 
         //                                     m_pestCurrent.dSlope, 
         //                                     m_pestCurrent.dProbit,
         //                                     m_pestCurrent.dSafetyFactor );

		 // we should really be using the planned exposure time, but since
		 // plan is always = actual then we can just use projected exposure time
		 double dExposureTime = odtsProjectedExposureTime.GetTotalSeconds() / 3600;

		 double dTargetDose = GetTargetDose (pCalcInfo->dInitConc,
			                                 pCalcInfo->dPestNValue,
											 pCalcInfo->dPlannedHLT,
											 dExposureTime);


         pCalcInfo->dDeltaDose = GetDeltaDose( dTargetDose, pCalcInfo->dAchievedDose );

         // Use the alternate calculation for the case PP_PLATEAUED
         COleDateTimeSpan odtsCurrentTime = odtCurrentTime - pConcentrationArray->GetAt( 0 ).odtDateTime;
         CONCENTRATION_DATA lastMaxConcentration = pConcentrationArray->GetAt( nLastMaxIndex );
         
		 //JNM 04/16/03 - Pass new argument pest n value
		 pCalcInfo->dProjectedFumigationTime = GetProjectedFumigationTime2( dCurrentConcentration, 
                                                                            pCalcInfo->dDeltaDose,
																			m_pestCurrent.dNValue,
                                                                            odtsExposureTime );

         // Part II, Step 9
         // When determining if more product needs to be added, we will assume a 1 hour lag time
         // from the most recent monitoring time

         // Part II, Step 9B
         // Use the alternate calculation for the case PP_PLATEAUED:
         //    CN = Cnow
         pCalcInfo->dFutureC = dCurrentConcentration;

         // Part II, Step 9C
         pCalcInfo->dFutureDose = GetFutureDose( pCalcInfo->dFutureC, dCurrentConcentration );

         // Part II, Step 9D
         pCalcInfo->dRemainingDose = GetRemainingRequiredDose( dTargetDose, 
                                                               pCalcInfo->dAchievedDose, 
                                                               pCalcInfo->dFutureDose );

         // Part II, Step 9E
         // Use the alternate calculation for the case PP_PLATEAUED
         pCalcInfo->dNewTargetC = GetNewTargetConcentration2( pCalcInfo->dRemainingDose,
                                                              odtsExposureTime, 
                                                              odtsProjectedExposureTime );
      }
      else // peakPattern == PP_INCREASING
      {
         // Set the remaining values to -1.0;
         // Fill up the return structure.
         pCalcInfo->dActualOHLT = -1.0;
         pCalcInfo->dActualEHLT = -1.0;
         pCalcInfo->dRemainingCT = -1.0;
         pCalcInfo->dProjectedCT = -1.0;
         pCalcInfo->dDeltaDose = -1.0;
         pCalcInfo->dProjectedFumigationTime = -1.0;
         pCalcInfo->dFutureC = -1.0;
         pCalcInfo->dFutureDose = -1.0;
         pCalcInfo->dRemainingDose = -1.0;
         pCalcInfo->dNewTargetC = -1.0;
      }

      bRet = TRUE;
   }

   return bRet;
}


double DMonitorPoint::GetCTAchieved( const CONCENTRATION_DATA_ARRAY *pdataConc )
{
   // The following is the function to determine the achieved CT
   //
   //    CTAchieved = 0.5 * sum[(Ci + C) * (T - Ti)]
   //
   // where Ci is the previous concentration,
   // Ti is the time at which the concentration Ci is read,
   // C is the current concentration,
   // and T is the current time, at which C is read.
   //
   // JNM 04/11/03 - Fumiguide Enh. 2003 - The calculation of the CT Achieved
   //		value will be changed such that the original trapezoidal equation
   //		will be used for increasing and plateau concentrations and the analytic 
   //		exponential integration will be used for decreasing concentrations.
   //		The total CTAchieved will therefore be the sum of these 2 equations.
   //       The formula for the analytic exponential integration is:
   //
   //		CTAchieved = Ci * HLT * [1 - pow(2,(Ti - T) / HLT)] / ln2
   //         
   //         where Ci is the previous concentration
   //         HLT is the half-loss time from the previous pt to the current pt
   //         Ti is the time at which the concentration Ci is read
   //         and T is the current time
   //

   double dConc = 0.0;
   COleDateTimeSpan odtspan = 0.0;
   double dCTSum = 0.0;
   CONCENTRATION_DATA dataConcInit;
   CONCENTRATION_DATA dataConcFinal;

   int nTotal = pdataConc->GetSize();

   if( nTotal == 1 )
   {
      // If there is only 1 concentration/time pair, then this is the start of exposure and our sum is 0
      dCTSum = 0;
   }
   else
   if( nTotal > 1 )
   {
      for( int nIndex = 1; nIndex < nTotal; nIndex++ )
      {
         dataConcInit = pdataConc->GetAt(nIndex - 1);
         dataConcFinal = pdataConc->GetAt(nIndex);
      
         odtspan = dataConcFinal.odtDateTime - dataConcInit.odtDateTime;
		 
		 // JNM 04/11/03 - Fumiguide Enh. 2003 - Determine if the previous 
		 //		concentration is greater than the current concentration.
		 // JNM 04/23/03 - AT Fix - Consider also the case wherein conc. have
		 //		the same time.
		 if(dataConcInit.dConcentration > dataConcFinal.dConcentration && odtspan.m_span > 0)
		 {
			// JNM 04/11/03 - Use analytic exponential integration since
			//		concentration is decreasing.

			// First get the HLT using the formula:
			//               ln(2) * (T - Tmax)
			// ActualOHLT  = ------------------
			//                ln(Cmax) - ln(C)
			//
			// where T is the current time, C is the current concentration,
			// Cmax is maximum concentration since last increase in concentration,
			// and Tmax is the time at which Cmax was observed.

			if( dataConcInit.dConcentration <= 0.0 )
			{
			  // Can't take the log of 0 or a negative number, therefore 
			  // we have to set the concentration value close to 0.
			  dataConcInit.dConcentration = 0.001;
			}

			if( dataConcFinal.dConcentration <= 0.0 )
			{
			  // Can't take the log of 0 or a negative number, therefore 
			  // we have to set the concentration value close to 0.
			  dataConcFinal.dConcentration = 0.001;
			}

			double dNumerator = ::log( 2 ) * ( odtspan.GetTotalSeconds() / 3600 ); // 3600 seconds/hour
			double dDenominator = ::log( dataConcInit.dConcentration ) - ::log( dataConcFinal.dConcentration );

			if( dDenominator == 0 )
			{
			  // Can't divide by 0, therefore proceed to the next iteration
			  continue;
			}

			double dHLT = dNumerator / dDenominator;

			//Second, compute CT achieved using:
			//	CTAchieved = Ci * HLT * [1 - pow(2,(Ti - T) / HLT)] / ln2

			double dBracket = 1 - ::pow(2,(-(odtspan.GetTotalSeconds() / 3600) / dHLT));

			dCTSum += dataConcInit.dConcentration * dHLT * dBracket / ::log(2);

		 }
		 
		 else
		 {
			 //We will use trapezoidal approximation since the concentration is
			 //increasing.

			 dConc = dataConcInit.dConcentration + dataConcFinal.dConcentration;

			 dCTSum += 0.5 * dConc * ( odtspan.GetTotalSeconds() / 3600 ); // 3600 seconds/hour
      
		 }
	  }   
   }
   else
   {
      dCTSum = -1;
   }

   return dCTSum;
}


double DMonitorPoint::GetDoseAchieved(const CONCENTRATION_DATA_ARRAY *pdataConc)
{
   // The following is the function to determine the achieved CT
   //
   //                              n    n
   //    CTAchieved = 0.5 * sum[(Ci  + C ) * (T - Ti)]
   //
   // where Ci is the previous concentration,
   // Ti is the time at which the concentration Ci is read,
   // C is the current concentration,
   // and T is the current time, at which C is read,
   // and n is the current N value for the pest/temperature combination.
   //
   // JNM 04/11/03 - Fumiguide Enh. 2003 - The calculation of the CT Achieved
   //		value will be changed such that the original trapezoidal equation
   //		will be used for increasing and plateau concentrations and the analytic 
   //		exponential integration will be used for decreasing concentrations.
   //		The total CTAchieved will therefore be the sum of these 2 equations.
   //       The formula for the analytic exponential integration is:
   //
   //		CTAchieved = pow(Ci, n) * HLT * [1 - pow(2,n * (Ti - T) / HLT)] / (n * ln2)
   //         
   //         where Ci is the previous concentration
   //         HLT is the half-loss time from the previous pt to the current pt
   //         Ti is the time at which the concentration Ci is read
   //         and T is the current time
   //		  and n is the current N value for the pest/temperature combination
   //

   double dConc = 0.0;
   COleDateTimeSpan odtspan = 0.0;
   double dDoseSum = 0.0;
   CONCENTRATION_DATA dataConcInit;
   CONCENTRATION_DATA dataConcFinal;
   double n = this->m_pestCurrent.dNValue;

   int nTotal = pdataConc->GetSize();

   if( nTotal == 1 )
   {
      // If there is only 1 concentration/time pair, then this is the start of exposure and our sum is 0
      dDoseSum = 0;
   }
   else
   if( nTotal > 1 )
   {
      for(int nIndex = 1; nIndex < nTotal; nIndex++)
      {
         dataConcInit = pdataConc->GetAt(nIndex - 1);
         dataConcFinal = pdataConc->GetAt(nIndex);
      
         odtspan = dataConcFinal.odtDateTime - dataConcInit.odtDateTime;

		 // JNM 04/11/03 - Fumiguide Enh. 2003 - Determine if the previous 
		 //		concentration is greater than the current concentration.
		 // JNM 04/23/03 - AT Fix - Consider also the case wherein conc. have
		 //		the same time.
		 if(dataConcInit.dConcentration > dataConcFinal.dConcentration && odtspan.m_span > 0)
		 {
			// JNM 04/11/03 - Use analytic exponential integration since
			//		concentration is decreasing.

			// First get the HLT using the formula:
			//               ln(2) * (T - Tmax)
			// ActualOHLT  = ------------------
			//                ln(Cmax) - ln(C)
			//
			// where T is the current time, C is the current concentration,
			// Cmax is maximum concentration since last increase in concentration,
			// and Tmax is the time at which Cmax was observed.

			if( dataConcInit.dConcentration <= 0.0 )
			{
			  // Can't take the log of 0 or a negative number, therefore 
			  // we have to set the concentration value close to 0.
			  dataConcInit.dConcentration = 0.001;
			}

			if( dataConcFinal.dConcentration <= 0.0 )
			{
			  // Can't take the log of 0 or a negative number, therefore 
			  // we have to set the concentration value close to 0.
			  dataConcFinal.dConcentration = 0.001;
			}

			double dNumerator = ::log( 2 ) * ( odtspan.GetTotalSeconds() / 3600 ); // 3600 seconds/hour
			double dDenominator = ::log( dataConcInit.dConcentration ) - ::log( dataConcFinal.dConcentration );

			if( dDenominator == 0 )
			{
			  // Can't divide by 0, therefore proceed to the next iteration
			  continue;
			}

			double dHLT = dNumerator / dDenominator;

			//Second, compute CT achieved using:
			//	CTAchieved = Ci * HLT * [1 - pow(2, n * (Ti - T) / HLT)] / (n * ln2)

			double dBracket = 1 - ::pow(2, n * (-(odtspan.GetTotalSeconds() / 3600) / dHLT));

			dDoseSum += ::pow(dataConcInit.dConcentration, n) * dHLT * dBracket / (n * ::log(2));

		 }
		 
		 else
		 {
			 //We will use trapezoidal approximation since the concentration is
			 //increasing.

			 if(  dataConcInit.dConcentration < 0.0 || dataConcFinal.dConcentration < 0.0 )
			 {
				// Can't take the power of a negative number
				return -1;
			 }
			 else
			 {
				dConc = ::pow( dataConcInit.dConcentration, n ) + ::pow( dataConcFinal.dConcentration, n );
			 }

			 dDoseSum += 0.5 * dConc * (odtspan.GetTotalSeconds() / 3600); // 3600 seconds/hour
		 }
	  }
   }
   else
   {
      dDoseSum = -1;
   }

   return dDoseSum;
}


PEAK_PATTERN DMonitorPoint::GetPeakPattern(const CONCENTRATION_DATA_ARRAY *padataConc, 
                                           const int nPeakIndex )
{
   int nLastReadingIndex = padataConc->GetUpperBound();

   // case 1: The current concentration reading is increasing, so there is no HLT yet.
   if ( nLastReadingIndex == nPeakIndex )
   {
      return PP_INCREASING;
   }

   // case 2: The current concentration reading is equal to the previous peak's
   //         reading, so the HLT not measurable - effectively infinite.
   if ( padataConc->GetAt(nLastReadingIndex).dConcentration == 
        padataConc->GetAt(nPeakIndex).dConcentration )
   {
      return PP_PLATEAUED;
   }

   // case 3: The current concentration reading is less than a previous peak, so the
   //         HLT can be calculated following Part II, Step 2 equations.
   return PP_DECREASING;
}


double DMonitorPoint::GetActualOHLT(const CONCENTRATION_DATA_ARRAY *padataConc, 
                                    const int nInitIndex,
                                    COleDateTime &odtHltDateTime)
{
   // The following is the function to determine the estimated HLT

   //               ln(2) * (T - Tmax)
   // ActualOHLT  = ------------------
   //                ln(Cmax) - ln(C)
   //
   // where T is the current time, C is the current concentration,
   // Cmax is maximum concentration since last increase in concentration,
   // and Tmax is the time at which Cmax was observed.

   CONCENTRATION_DATA dataConcInit = padataConc->GetAt( nInitIndex );
   CONCENTRATION_DATA dataConcLast = padataConc->GetAt( padataConc->GetUpperBound() );

   odtHltDateTime = dataConcLast.odtDateTime;

   COleDateTimeSpan odtspan = dataConcLast.odtDateTime - dataConcInit.odtDateTime;

   if( dataConcInit.dConcentration <= 0.0 || dataConcLast.dConcentration <= 0.0 )
   {
      // Can't take the log of 0 or a negative number
	  if (dataConcLast.dConcentration == 0.0)
		  dataConcLast.dConcentration = 0.001;
	  else
		  return -1;

   }

   double dNumerator = ::log( 2 ) * ( odtspan.GetTotalSeconds() / 3600 ); // 3600 seconds/hour
   double dDenominator = ::log( dataConcInit.dConcentration ) - ::log( dataConcLast.dConcentration );

   if( dDenominator == 0 )
   {
      return -1.0;
   }

   return( dNumerator / dDenominator );
}


double DMonitorPoint::GetRemainingCT( const double& dCurrentConcentration,
                                      const COleDateTimeSpan& odtsCurrentTime,
                                      const COleDateTimeSpan& odtsTotalFumigationTime,
                                      const double& dActualOHLT )
{
   // Equation for RemainingD, per Part II, Step 3, on 2001-07-25
   //
   //
   //                 C * ActualHLT           (T - Tf)/ActualHLT
   // CTRemainingD = --------------- * [ 1 - 2                   ]
   //                     ln(2) 
   //
   // where C is the current concentration reading,
   // ActualHLT is the actual HLT up to the current time,
   // T is the current time,
   // and Tf is the exepcted final time.

   double C  = dCurrentConcentration;
   double T  = odtsCurrentTime.GetTotalSeconds() / 3600;
   double Tf = odtsTotalFumigationTime.GetTotalSeconds() / 3600;

   double rightSide = 1 - pow( 2, ( T - Tf ) / dActualOHLT );
   double leftSide  = ( C * dActualOHLT ) / log( 2 );

   return( leftSide * rightSide );
}


double DMonitorPoint::GetRemainingCT2( const double& dCurrentConcentration,
                                       const COleDateTimeSpan& odtsCurrentTime,
                                       const COleDateTimeSpan& odtsTotalFumigationTime )
{
   // Alternate Equation for RemainingD, per Part II, Step 3, on 2001-12-21
   //
   // CTRemainingD = C * (Tf - T)
   //
   // where C is the current concentration reading,
   // T is the current time,
   // and Tf is the expected final time.

   double T  = odtsCurrentTime.GetTotalSeconds() / 3600;
   double Tf = odtsTotalFumigationTime.GetTotalSeconds() / 3600;

   return( dCurrentConcentration * (Tf - T) );
}


double DMonitorPoint::GetProjectedCT(const double& dAchievedCT, const double& dRemainingCT)
{
   // Equation for ProjectedD, per Part II, Step 4, on 2001-07-25
   //
   //    ProjectedD = CTAchievedD + CTRemainingD
   //
   // where CTAchievedD is the accumulated dosage from GetCTAchieved()
   // and RemainingD is the remaining dosage from GetRemainingDose().

   return( dAchievedCT + dRemainingCT );
}


double DMonitorPoint::GetDeltaDose(const double& dTargetDose, const double& dAchievedDose)
{
   // Equation for DeltaD, per Part II, Step 7, on 2001-07-25
   //
   //    DeltaD = CnTTargetDosage - CnTAchievedDosage
   //
   // where CnTTargetDosage is the target dose with the N factor, as per ::GetFinalDose().
   // and CnTAchievedDosage is the current accumlated dose with the N facto, as per GetDoseAchieved().

   return( dTargetDose - dAchievedDose );
}


//JNM 04/16/03 - Fum. Enh. 2003 - Added new argument dPestNValue to hold n value.
//		This is intended to hold the interpolated n value.
double DMonitorPoint::GetProjectedFumigationTime( const double& dActualEHLT,
                                                  const double& dCurrentConcentration,
                                                  const double& dDeltaDose,
												  const double& dPestNValue, 
                                                  const COleDateTimeSpan& odtsCurrentTime )
{
   // Equation for ExtraTime, per Part II Step 7, on 2001-11-20
   //
   //                   ActualEHLT              DeltaD * ln(2)
   //    ExtraTime = - ------------ * ln[1 - --------------------]
   //                     ln(2)                  n
   //                                        Cnow  * ActualEHLT
   //               |             |      |                       |
   //               +---chunk1----+      +--------chunk2---------+ 
   //
   // where ActualEHLT is the actual EHLT, given by GetEHLT();
   // DeltaD is the delta from the target dose, given by GetDeltaDose();
   // n is the n value associated with the current pest/temperature combination;
   // and Cnow is the most recent concentration observation.
   //
   //
   // Equation for NewFumeTime, per Part II Step 7, on 2001-12-20
   //
   //    NewFumeTime = Tnow + ExtraTime
   //
   // where Tnow is the time when the most recent concentration observation was taken;
   // and ExtraTime is from the equation above.

   if( dCurrentConcentration < 0.0 )
   {
      // Can't take a power of a negative number
      return -1;
   }

   double dChunk1 = -dActualEHLT / log( 2 );

   double dChunk2Num = dDeltaDose * log( 2 );

   //JNM 04/16/03 - Fum. Enh. 2003 - Instead of getting n from m_pestCurrent,
   //	this value is obtained from the passed parameter.
   //double dChunk2Den = pow( dCurrentConcentration, m_pestCurrent.dNValue ) * dActualEHLT;
   double dChunk2Den = pow( dCurrentConcentration, dPestNValue ) * dActualEHLT;

   if( dChunk2Den == 0.0 )
   {
      // Don't want to have a 0 in the denominator
      return -1;
   }

   double dChunk2 = 1 - ( dChunk2Num / dChunk2Den );
   
   if( dChunk2 <= 0.0 )
   {
      // Negative numbers don't play well with natural logs
      // Will return -1 and check up in DoMonitorPointCalculations()
      return -1;
   }

   double dExtraTime = dChunk1 * log( dChunk2 );

   // Need the amount of time since exposure, in hours.
   double dTNow = odtsCurrentTime.GetTotalSeconds() / 3600;  // 3600 sec/hr

   return( dTNow + dExtraTime );
}


//JNM 04/16/03 - Fum. Enh. 2003 - Added new argument to hold n pest value
double DMonitorPoint::GetProjectedFumigationTime2( const double& dCurrentConcentration,
                                                   const double& dDeltaDose,
												   const double& dPestNValue, 
                                                   const COleDateTimeSpan& odtsCurrentTime )
{
   // Alternate Equation for ExtraTime, per Part II Step 7, on 2001-12-21
   //
   //                  DeltaD 
   //    ExtraTime =  -------- 
   //                       n
   //                   Cnow 
   //
   // where DeltaD is the delta from the target dose, given by GetDeltaDose();
   // n is the n value associated with the current pest/temperature combination;
   // and Cnow is the most recent concentration observation.
   //
   //
   // Equation for NewFumeTime, per Part II Step 7, on 2001-12-20
   //
   //    NewFumeTime = Tnow + ExtraTime
   //
   // where Tnow is the time when the most recent concentration observation was taken;
   // and ExtraTime is from the equation above.

   if( dCurrentConcentration < 0.0 )
   {
      // Can't take a power of a negative number
      return -1;
   }

   //JNM 04/16/03 - Fum. Enh. 2003 - Get the n value from the passed argument
   //double dDenominator = pow( dCurrentConcentration, m_pestCurrent.dNValue );
   double dDenominator = pow( dCurrentConcentration, dPestNValue );


   if( dDenominator == 0.0 )
   {
      // Don't want to have a 0 in the denominator
      return -1;
   }

   double dExtraTime = dDeltaDose / dDenominator;

   // Need the amount of time since exposure, in hours.
   double dTNow = odtsCurrentTime.GetTotalSeconds() / 3600;  // 3600 sec/hr

   return( dTNow + dExtraTime );
}


double DMonitorPoint::GetFutureConcentration( const double& dCurrentConcentration, const double& dActualHLT )
{
   // Equation for future concentration 1 hour from the current time, per Part II, Step 9B, on 2001-11-20.
   //
   //                 (-DELTA_TIME_A / ActualOHLT )
   //    CN = Cnow * 2
   //
   // where Cnow is the current concentration
   // ActualHLT current HLT between Tmax and T,
   
	// CG 08/28/03 Changed the exponent numerator from 1 to DELTA_TIME_A
	//return( dCurrentConcentration * ::pow( 2, -1 / dActualHLT ) );
   return( dCurrentConcentration * ::pow( 2, -DELTA_TIME_A / dActualHLT ) );
}


double DMonitorPoint::GetFutureDose( const double& dFutureConcentration, const double& dCurrentConcentration )
{
   // Equation for future dose, per Part II, Step 9C, on 2001-12-18.
   //
   //                  n      n
   //                CN + Cnow
   //    NDosage = (------------) DELTA_TIME_A
   //                   2
   //
   // where CN is the future concentration,
   //       Cnow is the current concentration.
   //       Tlast is the time of the last monitoring point reading,
   //       T is the current time since start of exposure.
   
   if( dFutureConcentration < 0.0 || dCurrentConcentration < 0.0 )
   {
      // Can't take the power of a negative number
      return -1;
   }

   double dChunk1 = ::pow( dFutureConcentration, this->m_pestCurrent.dNValue );
   double dChunk2 = ::pow( dCurrentConcentration, this->m_pestCurrent.dNValue );

   // CG 08/28/03 Added the multiplier DELTA_TIME_A
   //return( ( dChunk1 + dChunk2 ) / 2.0);
   return( ( ( dChunk1 + dChunk2 ) / 2.0) * DELTA_TIME_A );
}


double DMonitorPoint::GetRemainingRequiredDose( const double& dTargetDose,
                                                const double& dAchievedDose,
                                                const double& dFutureDose )
{
   // Equation for Remaining Required Dose, per Part II, Step 9D, on 2001-11-20
   //
   //    DeltaD = CnTTargetD - (CnTAchievedD + NDosage)
   //
   // where CnTTargetD is the target dose,
   // CntAchievedD is the achived dose,
   // and NDosage is the future dose.

   return( dTargetDose - ( dAchievedDose + dFutureDose ) );
}


double DMonitorPoint::GetNewTargetConcentration( const double& dRemainingReqDose, 
                                                 const double& dActualEHLT,
                                                 const COleDateTimeSpan& odtsCurrentTime,
                                                 const COleDateTimeSpan& odtsTotalFumigationTime )
{
   // Equation for New Target Concentration, per Part II, Step 9E, on 2002-01-15
   //
   //                      DeltaD                                 1/n
   // NTC = { -------------------------------------------------- } 
   //          ActualEHLT           -(Tf - Tnow - DELTA_TIME_A)/ActualEHLT
   //         ------------ * [1 - 2                            ]
   //            ln(2)
   //
   //         |--chunk1--|   |-------------chunk2--------------|
   //         |---------------------chunk3---------------------|
   //
   // where DeltaD is the remaining required dose,
   // ActualEHLT is the current HLT between Tmax and T,
   // Tf is the total fumigation duration,
   // and Tnow is the current fumigation duration.
   //
   // JNM 05/12/03 - PT fix - New Target Concentration does not take into
   //     account that you are achieving. Formula 9D has no addition
   //     equilibrium time. Therefore, as a fix, the "1 hr" used in the
   //     formula will be shortened to 30 minutes.
   //        Tf - Tnow - 0.5
   //
   // CG 08/27/03 - Changed the "1hr" used in the formula to DELTA_TIME_A
   // which represents the amt of time that it takes to reach a new equilibrium 
   // concentration once additional gas has been injected

   double dChunk1 = dActualEHLT / ::log( 2 );

   // Need the fumigation times in hours (3600 seconds/hour)
   double dTf = odtsTotalFumigationTime.GetTotalSeconds() / 3600;
   double dTnow = odtsCurrentTime.GetTotalSeconds() / 3600;
   
   //JNM 05/12/03 - Change 1hr to 30 minutes
   //double dChunk2 = 1 - ::pow( 2, ( -(dTf - dTnow - 1) / dActualEHLT) );
   //double dChunk2 = 1 - ::pow( 2, ( -(dTf - dTnow - 0.5) / dActualEHLT) );
   
   // CG 08/27/03 - Changed the "1hr" used in the formula to DELTA_TIME_A
   double dChunk2 = 1 - ::pow( 2, ( -(dTf - dTnow - DELTA_TIME_A) / dActualEHLT) );

   double dDenominator = dChunk1 * dChunk2;

   // Let's do some checking before we calc the final answer
   if( dDenominator == 0.0 )
   {
      // Can't divide by 0
      return -1;
   }

   double dChunk3 = dRemainingReqDose / dDenominator;

   if( dChunk3 < 0.0 )
   {
      // Can't take the power of a negative number
      return -1;
   }

   return( ::pow( dChunk3, 1 / m_pestCurrent.dNValue ) );
}


double DMonitorPoint::GetNewTargetConcentration2( const double& dRemainingReqDose, 
                                                  const COleDateTimeSpan& odtsCurrentTime,
                                                  const COleDateTimeSpan& odtsTotalFumigationTime  )
{
   // Alternate Equation for New Target Concentration, per Part II, Step 9E, on 2002-01-15
   //
   //             DeltaD        1/n
   // NTC = { --------------- }
   //          Tf - Tnow - 1
   //
   // where DeltaD is the remaining required dose,
   // Tf is the total fumigation duration,
   // and Tnow is the current fumigation duration.

   // Need the fumigation times in hours (3600 seconds/hour)
   //
   // JNM 05/12/03 - PT fix - New Target Concentration does not take into
   //     account that you are achieving. Formula 9D has no addition
   //     equilibrium time. Therefore, as a fix, the "1 hr" used in the
   //     formula will be shortened to 30 minutes.
   //        Tf - Tnow - 0.5

   // CG 08/27/03 - Changed the "1hr" used in the formula to DELTA_TIME_A
   // which represents the amt of time that it takes to reach a new equilibrium 
   // concentration once additional gas has been injected

   double dTf = odtsTotalFumigationTime.GetTotalSeconds() / 3600;
   double dTnow = odtsCurrentTime.GetTotalSeconds() / 3600;
   
   double dChunk = 0.0;

   // JNM 05/12/03 - PT fix - Change 1 hr to 30 minutes
   //double dDenominator = dTf - dTnow - 1;

   // CG 08/27/03 - Changed the "1hr" used in the formula to DELTA_TIME_A
   //double dDenominator = dTf - dTnow - 0.5;
   double dDenominator = dTf - dTnow - DELTA_TIME_A;


   if( dDenominator == 0.0 )
   {
      // Can't divide by 0
      return -1;
   }

   dChunk = dRemainingReqDose / dDenominator;

   // Let's do some checking before we calc the final answer

   if( dChunk < 0.0 )
   {
      // Can't take the power of a negative number
      return -1;
   }

   return( ::pow( dChunk, 1 / m_pestCurrent.dNValue ) );
}


int DMonitorPoint::GetLastPeakIndex( const CONCENTRATION_DATA_ARRAY *padataConc,
                                     COleDateTime &odtLastPeakDateTime )
{
   // We will work our way from the bottom of the array to the top
   // When we find the first drop in value then we know we have achieved the last peak

   int nMaxIndex = padataConc->GetUpperBound();

   CONCENTRATION_DATA dataConc;

   int nIndexPeak = -1;
   double dPeakConc = -1.0;

   for( int nIndex = nMaxIndex; nIndex >= 0; nIndex-- )
   {
      dataConc = padataConc->GetAt( nIndex );

      if( dataConc.dConcentration >= dPeakConc )
      {
         dPeakConc = dataConc.dConcentration;
         odtLastPeakDateTime = dataConc.odtDateTime;

         nIndexPeak = nIndex;
      }
      else
      {
         // Break out of the loop, we have found our last peak value
         break;
      }
   }

   return nIndexPeak;
}


BOOL DMonitorPoint::GetPt2PtHlt( LPMONITOR_POINT_2_POINT_DATA lpdataMonPt2Pt )
{
   ASSERT( lpdataMonPt2Pt );

   // Get the last peak concentration from the concentration array
   int nLastPeakIndex = GetLastPeakIndex( lpdataMonPt2Pt->padataConc, lpdataMonPt2Pt->odtLastPeakDateTime );

   // Get the Peak Pattern
   PEAK_PATTERN ePeakPattern = GetPeakPattern( lpdataMonPt2Pt->padataConc, nLastPeakIndex );

   switch( ePeakPattern )
   {
   case PP_DECREASING:
      lpdataMonPt2Pt->dActualOHLT = GetActualOHLT( lpdataMonPt2Pt->padataConc, nLastPeakIndex, lpdataMonPt2Pt->odtHltDateTime );
      break;

   case PP_PLATEAUED:
      lpdataMonPt2Pt->dActualOHLT = 0.0;
      break;

   case PP_INCREASING:
      // We can't calc an HLT if concentrations are increasing or not changing (infinity)
      lpdataMonPt2Pt->dActualOHLT = -1;
      break;

   default:
      ASSERT( FALSE );
   }

   return TRUE;
}


double DMonitorPoint::GetTargetDose(const double dInitialConcentration, const double dnValue, const double dOHLT, const double dExposureTime)
{

   // The following is the function to determine CnTTargetD
   //                   n
   //          HLTe * Co           -(t / HLTe)
   // Dose  = ---------- * [ 1 - 2            ]
   //            ln(2)
   //
   // Co = dInitialConcentration
   // HLTe = dHltEstimated
   // t = dExposuretime
   // n = dnValue
   //
   // 

   //                 n
   //CnTTargetD =   Co   *    HLTe                ((-t * n) / HLTe)
   //                       ---------  *  [  1 - 2                  ]
   //                       n * ln(2)


	// (C0calc^nhi)*(tau/(nhi*LN(2)))*(1 - 2^(-tf*nhi/tau))

   // EHLT = OHLT / n value
   //double dHltEstimated = dOHLT / dnValue;
   double dHltEstimated = dOHLT;

   double d2RaisedPower = pow( 2, - ( dExposureTime*dnValue / dHltEstimated ) );

   double dInitConcRaisedPower = pow( dInitialConcentration, dnValue );

   double Group2 = dHltEstimated / (dnValue*log(2));

   return dInitConcRaisedPower * Group2 * (1 - d2RaisedPower);
   
   //return( ( ( dHltEstimated * dInitConcRaisedPower ) / dnValue * log( 2 ) ) * ( 1 - d2RaisedPower ) );

}
