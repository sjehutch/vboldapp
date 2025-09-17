// Area.cpp: implementation of the DArea class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "Area.h"

#include "CommodityData.h"
#include "DasCommodityData.h"
#include "DasPestData.h"
#include "GlobalEnums.h"
#include "GlobalFunctions.h"
#include "PestData.h"



#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DArea::DArea()
{

}


DArea::~DArea()
{

}


BOOL DArea::GetAreaStatus( LPAREA_DATA lpAreaData )
{
   ASSERT( lpAreaData );

   if( lpAreaData->padataAreaMonPt->GetSize() == 0 )
   {
      // There are no monitor points to use 
      lpAreaData->dAchievedConcentrationTime = -1.0;
      lpAreaData->dAchievedDose = -1.0;
      lpAreaData->dAddProfume = -1.0;
      lpAreaData->dHlt = -1.0;
      lpAreaData->dProjectedConcentrationTime = -1.0;
      lpAreaData->dProjectedTime = -1.0;
      lpAreaData->dMeanNTC = -1.0;
      lpAreaData->nStatus = STATUS_UNKNOWN;
   }

   BOOL bReturn = TRUE;

   double dMeanFutureConcentration = 0.0;
   double dElapsedTime = 0.0;

   // Initialize the status
   lpAreaData->nStatus = STATUS_UNKNOWN;

   // Get the basic mean information
   GetMeanValues( lpAreaData->dAchievedConcentrationTime,
                  lpAreaData->dAchievedDose,
                  lpAreaData->dHlt,
                  lpAreaData->dProjectedConcentrationTime,
                  lpAreaData->dProjectedTime,
                  lpAreaData->dMeanNTC,
                  dMeanFutureConcentration,
                  dElapsedTime,
                  lpAreaData->padataAreaMonPt );

   // No ProFume needs to be added yet
   lpAreaData->dAddProfume = -1.0;

   // Decision making time
   // Check if the achieved concentration-time is above the target concentration-time
   if( lpAreaData->dAchievedConcentrationTime >= lpAreaData->dTargetConcentrationTime )
   {
      lpAreaData->nStatus = STATUS_ACHIEVED_TARGET;

      return bReturn;
   }

   // Check if there is an HLT, otherwise we are increasing in concentration
   // An HLT of 0.0 is the plateau case
   if( lpAreaData->dHlt < 0.0 )
   {
      lpAreaData->nStatus = STATUS_INCREASING_CONCENTRATION;

      return bReturn;
   }

   // Get the maximum permitted concentration time based on the commodity
   // Since this is very short term for the function, will create it dynamically
   DDasCommodityData *pdasCommodityData = NULL;
   pdasCommodityData = new DDasCommodityData();

   pdasCommodityData->PopulateCommodityData( lpAreaData->eCommodity,
                                             lpAreaData->eFumigationType,
                                             lpAreaData->eTreatmentType,
                                             lpAreaData->strMsg );
   
   COMMODITY_DATA dataCommodity = pdasCommodityData->GetCommodityData();

   delete pdasCommodityData;
   pdasCommodityData = NULL;

   // Need to check the projected concentration-time, and see if it is on target (+-2% and <= Maximum Permitted Concentration Time (CT))
   double dMinProjConcTime = lpAreaData->dTargetConcentrationTime * ( 1 - TARGET_RANGE );
   double dMaxProjConcTime = min( lpAreaData->dTargetConcentrationTime * ( 1 + TARGET_RANGE ), dataCommodity.dMaxConcentrationTime );

   if( lpAreaData->dProjectedConcentrationTime >= dMinProjConcTime && 
       lpAreaData->dProjectedConcentrationTime <= dMaxProjConcTime )
   {
      // Nothing needs to be done
      lpAreaData->nStatus = STATUS_ON_TARGET;
   }
   else
   if( lpAreaData->dProjectedConcentrationTime < dMinProjConcTime )
   {
      lpAreaData->nStatus = STATUS_BELOW_TARGET;

      // Determine how much more time needs to be added
      // This has already been done 
      // The mean projected time has already been calculated in GetMeanValues()
   
      // Determine how much more ProFume needs to be added
      // Check the MeanNTC and ensure it does not exceed the maximum permitted concentration
      // This corresponds to Part II eqn 9G.
      if( lpAreaData->dMeanNTC > MAX_CONCENTRATION )
      {
         lpAreaData->dMeanNTC = MAX_CONCENTRATION;
         
         lpAreaData->nStatus = STATUS_TARGET_DOSE_INCOMPATABLE;
      }

      // 12/10/02 Jett Gamboa - added check for elapsed times
      // only calculate if ALL elapsed times are within longest elapsed time - 30 minutes
      if (CheckElapsedTimes(lpAreaData->padataAreaMonPt))
      {
         // This corresponds to Part II eqn 9H.
         lpAreaData->dAddProfume = AddProfume( lpAreaData->dVolume,
                                               lpAreaData->dMeanNTC,
                                               dMeanFutureConcentration );
      }
      else
      {
         lpAreaData->dAchievedConcentrationTime = -1;
         lpAreaData->dProjectedConcentrationTime = -1;
         lpAreaData->dMeanNTC = -1;

         lpAreaData->nStatus = STATUS_BAD_ELAPSED_TIMES;
      }
   }
   else
   if( lpAreaData->dProjectedConcentrationTime > dMaxProjConcTime )
   {
      // Nothing else needs to be done
      // The mean projected time has already been calculated in GetMeanValues()
      lpAreaData->nStatus = STATUS_ABOVE_TARGET;
   }
   else
   {
      // We should not be here
      ASSERT( FALSE );

      bReturn = FALSE;

      lpAreaData->nStatus = STATUS_UNKNOWN;
   }

   return bReturn;
}


void DArea::GetMeanValues( double &dAchievedConcentrationTime,
                           double &dAchievedDose,
                           double &dHlt,
                           double &dProjectedConcentrationTime,
                           double &dProjectedTime,
                           double &dFutureTargetConcentration,
                           double &dFutureConcentration,
                           double &dElapsedTime,
                           AREA_MON_PT_DATA_ARRAY* padataAreaMonPt )
{
   ASSERT( padataAreaMonPt );

   // Get the mean values of the parameters using the monitor point array.
   // Achieved concentration-time and achieved dose are based on the total number of monitor points
   // Calculations based on HLT are based on the number of monitor points that had an HLT value

   int nTotal = padataAreaMonPt->GetSize();

   if( nTotal == 0 )
   {
      // We should not be here
      ASSERT( FALSE );

      return;
   }

   int nTotalHlt = 0L;
   int nTotalProjectedTime = 0L;

   double dAchievedConcentrationTimeTotal = 0.0;
   double dAchievedDoseTotal = 0.0;
   double dHltTotal = 0.0;
   double dProjectedConcentrationTimeTotal = 0.0;
   double dProjectedTimeTotal = 0.0;
   double dFutureTargetConcentrationTotal = 0.0;
   double dFutureConcentrationTotal = 0.0;
   double dElapsedTimeTotal = 0.0;

   AREA_MON_PT_DATA dataAreaMonPt;

   // Get the total's
   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      dataAreaMonPt = padataAreaMonPt->GetAt( nIndex );
      
      dAchievedConcentrationTimeTotal += dataAreaMonPt.dAchievedConcentrationTime;
      dAchievedDoseTotal += dataAreaMonPt.dAchievedDose;
      
      if( dataAreaMonPt.dHlt >= 0.0 )
      {
         dHltTotal += dataAreaMonPt.dHlt;
         dProjectedConcentrationTimeTotal += dataAreaMonPt.dProjectedConcentrationTime;
         
         if( dataAreaMonPt.dProjectedTime >= 0.0 )
         {
            dProjectedTimeTotal += dataAreaMonPt.dProjectedTime;

            nTotalProjectedTime++;
         }
         dFutureTargetConcentrationTotal += dataAreaMonPt.dFutureTargetConcentration;
         dFutureConcentrationTotal += dataAreaMonPt.dFutureConcentration;
         dElapsedTimeTotal += dataAreaMonPt.dElapsedTime;

         nTotalHlt++;
      }
   }

   // Calc the mean value
   dAchievedConcentrationTime = dAchievedConcentrationTimeTotal / static_cast< double >( nTotal );
   dAchievedDose = dAchievedDoseTotal / static_cast< double >( nTotal );
   

   if( nTotalHlt > 0 )
   {
      dHlt = dHltTotal / static_cast< double >( nTotalHlt );
      dProjectedConcentrationTime = dProjectedConcentrationTimeTotal / static_cast< double >( nTotalHlt );
      
      if( nTotalProjectedTime > 0 )
      {
         dProjectedTime = dProjectedTimeTotal / static_cast< double >( nTotalProjectedTime );
      }
      else
      {
         dProjectedTime = -1;
      }

      dFutureTargetConcentration = dFutureTargetConcentrationTotal / static_cast< double >( nTotalHlt );
      dFutureConcentration = dFutureConcentrationTotal / static_cast< double >( nTotalHlt );
      dElapsedTime = dElapsedTimeTotal / static_cast< double >( nTotalHlt );
   }
   else
   {
      dHlt = -1;
      dProjectedConcentrationTime = -1;
      dProjectedTime = -1;
      dFutureTargetConcentration = -1;
      dFutureConcentration = -1;
      dElapsedTime = -1;
   }
}


double DArea::AddProfume( const double &dVolume,
                          const double &dFutureTargetConcentration,
                          const double &dFutureConcentration )
{
   // Compare the future target concentration to the max concentration value
   double dFutureTargetConcentrationMax = __min( dFutureTargetConcentration, MAX_CONCENTRATION );
   
   // Determine the amount of ProFume to add (kg)
   // kg = g/m^3 * m^3 / 1000 g/kg
   return ( dFutureTargetConcentrationMax - dFutureConcentration ) * dVolume / 1000;
}



BOOL DArea::CheckElapsedTimes(AREA_MON_PT_DATA_ARRAY* padataAreaMonPt)
{

	// 12/10/02 Jett Gamboa - added
	// check elapsed times of monitoring points belonging to the area. If elapsed time
   // for monitoring point within the area is more than 30 minutes less than the longest
   // elapsed time, return FALSE so we do not calculate 

   int nTotal = padataAreaMonPt->GetSize();
   double nMaxElapsedTime  = 0;
   int nIndex;

   if( nTotal == 0 )
   {
      // We should not be here
      ASSERT( FALSE );

      return TRUE;
   }

   AREA_MON_PT_DATA dataAreaMonPt;

   // loop through all the Data points and get the maximum elapsed time
   for( nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      dataAreaMonPt = padataAreaMonPt->GetAt( nIndex );

      // set the maximum elapsed time if we have one so far
      if(dataAreaMonPt.dElapsedTime > nMaxElapsedTime)
      {
         nMaxElapsedTime = dataAreaMonPt.dElapsedTime;
      }
   }

   // return true if there was no elapsed time or if elapsed time is not yet
   // over 30 minutes
   if(nMaxElapsedTime == 0 || ((nMaxElapsedTime - ELAPSED_TIME_DIFF) <= 0))
   {
      return TRUE;
   }

   // loop through the monitoring points again, this time, check if the elapsed time
   // is less than the maximum elapsed time - 30 minutes (ELAPSED_TIME_DIFF)
   for( nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      dataAreaMonPt = padataAreaMonPt->GetAt( nIndex );

      if(dataAreaMonPt.dElapsedTime < (nMaxElapsedTime - ELAPSED_TIME_DIFF))
      {
         return FALSE;
      }
   }

   return TRUE;
}
