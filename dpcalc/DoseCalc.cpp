// DoseCalc.cpp: implementation of the DDoseCalc class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "DoseCalc.h"
#include "GlobalEnums.h"
#include "GlobalFunctions.h"
#include <math.h>


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif



//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DDoseCalc::DDoseCalc()
{

}


DDoseCalc::~DDoseCalc()
{

}

int nCounterPrev = 0; // PAT 1/27/06 - Global Variable that will hold the Prev record's counter

/*
 ******************  The real calculations  *********************************************************
	Sub:    GetProfumeRequired
	Inputs: lpdataDose
	Desc:	Subroutine to set the required amound of fumigide 
			for the given pest under current conditions
			This function is called for the 
			CProFumeCalculator::GetProfumeRequired and receives the current pest object as an input
 *****************************************************************************************************
*/

BOOL DDoseCalc::GetProfumeRequired( LPDOSE_DATA lpdataDose )
{
   ASSERT( lpdataDose );

   BOOL bReturn = TRUE;

   double dFinalDoseHigh;
   double dFinalDoseLow;
   double dLoadFactor; //Pat 1/16/06 -  Added
   int IsComplete; //Pat 1/16/06 -  Added

   dLoadFactor = lpdataDose->dLoadFactor; 
   IsComplete = lpdataDose->intIsComplete; 

   // Populate the pest array with the correct pests information
   m_dasPestData.PopulatePestData( lpdataDose->ePest, 
                                   lpdataDose->eLifeStage, 
                                   lpdataDose->eTreatmentType,
                                   lpdataDose->eFumigationType,
                                   lpdataDose->dTimeExposure,
                                   lpdataDose->strError );


   // Get the current pest data for the high temp
   m_pestCurrent = m_dasPestData.GetPestData( lpdataDose->dTemp, FALSE );


   // Populate the commodity information
   m_dasCommodityData.PopulateCommodityData( lpdataDose->eCommodity, 
                                             lpdataDose->eFumigationType,                                       
                                             lpdataDose->eTreatmentType,
                                             lpdataDose->strError ); 
   
   // Get the commodity data based on the previous info
   m_dataCommodity = m_dasCommodityData.GetCommodityData();



   double dTempHigh = m_pestCurrent.dTemp;

   //12/23/04-YYT-If concentration time is > 0, then user has defined a CT. 
   //We will then skip steps 4 & 5. We just set the dFinalDoseHIgh to the user-defined CT.
   //else we do steps 4&5 and compute for Final Dose.
   //We define a boolean user_defined and set to true when there is a user-defined CT.
   //This is needed because dFinalDoseHigh can be Zero when GetFinalDose returns Zero.
   //We can not then use this dFinalDoseHigh as a gauge if there is a user-defined CT or none
   bool bUserDefined = FALSE;

   if ( lpdataDose->dConcentrationTime > 0 )
   {
		dFinalDoseHigh = lpdataDose->dConcentrationTime;
		bUserDefined = TRUE;
   }
   else
		// Part I Step 4 & 5, on 2002-07-10
		dFinalDoseHigh = ::GetFinalDose( m_pestCurrent.dIntercept, 
                                         m_pestCurrent.dSlope, 
                                         m_pestCurrent.dProbit,
                                         m_pestCurrent.dSafetyFactor );

   //12/22/04- YYT - Added bUserDefined boolean as parameter. Use 2 different formula
   //to get initial concentration, if user has a defined CT vs none.
   // Part I Step 6 & 7
   double dInitConcHigh = GetInitConcentration( lpdataDose->dHltOriginal, 
                                                lpdataDose->dTimeExposure,
                                                dFinalDoseHigh,
												bUserDefined,
												lpdataDose->dVolume,
											    lpdataDose->dTemp,												  
											    m_dataCommodity.dBConstant,
											    m_dataCommodity.dCConstant,
											    m_dataCommodity.dDConstant,
											    m_dataCommodity.dPorosity,
											    m_dataCommodity.dDCZero, 
											    dLoadFactor,
											    IsComplete,
												lpdataDose->intCounter);

    double dInitConcHigh_FORCT = GetInitConcentration( lpdataDose->dHltOriginal, 
                                                lpdataDose->dTimeExposure,
                                                dFinalDoseHigh,
												bUserDefined,
												lpdataDose->dVolume,
											    lpdataDose->dTemp,												  
											    m_dataCommodity.dBConstant,
											    m_dataCommodity.dCConstant,
											    m_dataCommodity.dDConstant,
											    m_dataCommodity.dPorosity,
											    m_dataCommodity.dDCZero, 
											    dLoadFactor,
											    0,
												lpdataDose->intCounter);

   // Get the current pest data for the low temp	m_pestCurrent.dIntercept	-6.9600000000000

   m_pestCurrent = m_dasPestData.GetPestData( lpdataDose->dTemp, TRUE );

   double dTempLow = m_pestCurrent.dTemp;

   // If the temps are the same, then we are working with the max
   // value to start with so don't do another calc, we already have
   // our answer
   if( dTempLow == dTempHigh )
   {
      lpdataDose->dInitialConcentration = dInitConcHigh;
	  lpdataDose->dInitialConcentration_FOR_CT = dInitConcHigh_FORCT;
      lpdataDose->dDose = dFinalDoseHigh;
   }
   else
   {

	  //12/23/04-YYT-do same checking here. if concentration time is >0 it means, user has defined a CT
	  //else we need to compute it here.. set the bUserDefined variable accordingly.

	  bUserDefined = FALSE;

	  if ( lpdataDose->dConcentrationTime > 0 )
	  {
		  dFinalDoseLow = lpdataDose->dConcentrationTime;
		  bUserDefined = TRUE;
	  }
	  else
			// Part I Step 4 & 5, on 2002-07-10
		   dFinalDoseLow = ::GetFinalDose( m_pestCurrent.dIntercept, 
                                             m_pestCurrent.dSlope,
                                             m_pestCurrent.dProbit,
                                             m_pestCurrent.dSafetyFactor );

      // Part I Step 6 & 7, on 2002-07-10
      // We need to get the concentration for the low value
	  //12/23/04-Added the flag bUserDefined.
      double dInitConcLow = GetInitConcentration( lpdataDose->dHltOriginal, 
                                                  lpdataDose->dTimeExposure,
                                                  dFinalDoseLow,
												  bUserDefined,
												  lpdataDose->dVolume,
												  lpdataDose->dTemp,												  
												  m_dataCommodity.dBConstant,
												  m_dataCommodity.dCConstant,
												  m_dataCommodity.dDConstant,
												  m_dataCommodity.dPorosity,
												  m_dataCommodity.dDCZero, 
												  dLoadFactor,
												  IsComplete,
												  lpdataDose->intCounter);

	  //07-04-07-David Smith Added Caclulation for CT corrections.
      double dInitConcLow_FORCT = GetInitConcentration( lpdataDose->dHltOriginal, 
                                                  lpdataDose->dTimeExposure,
                                                  dFinalDoseLow,
												  bUserDefined,
												  lpdataDose->dVolume,
												  lpdataDose->dTemp,												  
												  m_dataCommodity.dBConstant,
												  m_dataCommodity.dCConstant,
												  m_dataCommodity.dDConstant,
												  m_dataCommodity.dPorosity,
												  m_dataCommodity.dDCZero, 
												  dLoadFactor,
												  0,
												  lpdataDose->intCounter);
       

      // Need to do a straight line interpolation between the 2 values
      double dIntTemp = ( lpdataDose->dTemp - dTempLow ) / ( dTempHigh - dTempLow );

      
	  //One is needed for Dosage and one is needed for CT caclulation
	  lpdataDose->dInitialConcentration = (dInitConcHigh  - dInitConcLow ) * dIntTemp + dInitConcLow;
      



	  //07-04-07-David Smith Added Caclulation for CT corrections.
	  //One is needed for Dosage and one is needed for CT caclulation
	  lpdataDose->dInitialConcentration_FOR_CT = (dInitConcHigh_FORCT  - dInitConcLow_FORCT ) * dIntTemp + dInitConcLow_FORCT;



      // Caclculates the lpdataDose->dDose 
      lpdataDose->dDose                 = (dFinalDoseHigh - dFinalDoseLow) * dIntTemp + dFinalDoseLow;


   }

   // Warning handles both max concentration and max concentration-time warnings
   CString strConcentrationWarning;
   
   // Part I Step 8, on 2002-07-10
   // We need to check the interpolated concentration and see if it is greater
   // If it is then the user needs to be warned
   if( lpdataDose->dInitialConcentration > MAX_CONCENTRATION )
   {
      strConcentrationWarning.LoadString( IDS_WARNING_CO_EXCEEDED );
      
      lpdataDose->dInitialConcentration = MAX_CONCENTRATION;
   }

   //12/23/04-YYT- skip this step when user has defined CT
   // Get the concentration-time
   // Part I Step 9, on 2002-07-10
   if ( lpdataDose->dConcentrationTime == 0 )
		lpdataDose->dConcentrationTime = GetConcentrationTime( lpdataDose->dHltOriginal,
                                                          lpdataDose->dTimeExposure,
                                                          lpdataDose->dInitialConcentration_FOR_CT );

   // We need to check the interpolated concentration-time and see if it is greater
   // If it is then the user needs to be warned
   if( lpdataDose->dConcentrationTime > m_dataCommodity.dMaxConcentrationTime )
   {
      if( strConcentrationWarning.IsEmpty() )
      {
         strConcentrationWarning.Format( IDS_WARNING_CT_EXCEEDED, 
                                         m_dataCommodity.dMaxConcentrationTime,
                                         m_dataCommodity.dMaxConcentrationTime,
                                         m_dataCommodity.dMaxConcentrationTime );
      }
      else
      {
         strConcentrationWarning.Format( IDS_WARNING_CO_CT_EXCEEDED,
                                         m_dataCommodity.dMaxConcentrationTime,
                                         m_dataCommodity.dMaxConcentrationTime );
      }

      // Set the concentration-time to the max value, since we can't exceed it
      lpdataDose->dConcentrationTime = m_dataCommodity.dMaxConcentrationTime;

      // Part I Step 10B, on 2002-07-10
      lpdataDose->dInitialConcentration = GetMaxInitConcentration( m_dataCommodity.dMaxConcentrationTime,
                                                                   lpdataDose->dHltOriginal, 
                                                                   lpdataDose->dTimeExposure );

      // Get the current pest data for the high temp
      m_pestCurrent = m_dasPestData.GetPestData( lpdataDose->dTemp, FALSE );

      // Part I Step 10C, on 2002-07-10
      double dMaxDoseHigh = GetMaxDose( lpdataDose->dHltOriginal,
                                        lpdataDose->dTimeExposure, 
                                        lpdataDose->dInitialConcentration,
										IsComplete,
										bUserDefined);

      // Get the current pest data for the low temp
      m_pestCurrent = m_dasPestData.GetPestData( lpdataDose->dTemp, TRUE );

      double dTempLow = m_pestCurrent.dTemp;

      // If the temps are the same, then we are working with the max
      // value to start with so don't do another calc, we already have
      // our answer
      if( dTempLow == dTempHigh )
      {
         lpdataDose->dDose = dMaxDoseHigh;
      }
      else
      {
         // Part I Step 10C
         double dMaxDoseLow = GetMaxDose( lpdataDose->dHltOriginal,
                                          lpdataDose->dTimeExposure, 
                                          lpdataDose->dInitialConcentration,
										  IsComplete,
										  bUserDefined);

         double dIntTemp = ( lpdataDose->dTemp - dTempLow ) / ( dTempHigh - dTempLow );

         lpdataDose->dDose = ( dMaxDoseHigh - dMaxDoseLow ) * dIntTemp + dMaxDoseLow;
      }
   }


   // Place the warnings into the warning string
   if( !strConcentrationWarning.IsEmpty() )
   {
      if( !lpdataDose->strError.IsEmpty() )
      {
         lpdataDose->strError += _T( "\r\n" );
      }

      lpdataDose->strError += strConcentrationWarning;
   }


   // ProfumeRequired (kg) = IntConcentration (g/m^3) * 10^-3 (kg/g) * volume (m^3)

   // Consolidating the equation gives the following equation:

   // ProfumeRequired (kg) = IntConcentration (g/m^3) * volume (m^3) * 10^-3 (kg/g)

   // Part I Step 11
   //PAT 1/26/06 - If the criteria is complete and CT was not defined by the USER
   if ((IsComplete == 1) && ( bUserDefined == FALSE ) )
   {
		lpdataDose->dProfumeReq = lpdataDose->dInitialConcentration * GetCommodityDisplacement (lpdataDose->dVolume, dLoadFactor, m_dataCommodity.dPorosity ) * pow( 10, -3 ); 
   }
   else
   {
		lpdataDose->dProfumeReq = lpdataDose->dInitialConcentration * lpdataDose->dVolume * pow( 10, -3 ); 
   }

   // # of cylinders = ProfumeRequired (kg) / 56.7 (kg/cyl) 
   lpdataDose->dCylindersReq = lpdataDose->dProfumeReq / KILO_PER_CYL;


   return bReturn;
}


double DDoseCalc::GetEstimatedHlt( const double &dHltOriginal )
{
   return dHltOriginal / m_pestCurrent.dNValue;
}

double DDoseCalc::GetInitConcentration( const double &dHltOriginal, 
                                        const double &dTimeExposure, 
                                        const double &dFinalDose, 
										bool dUserDefined,
										const double &dVolume,
										const double &dTemp,
										const double &dBConstant,
										const double &dCConstant,
										const double &dDConstant,
										const double &dPorosity,
										const double &dDCZero,
										const double &dLoadFactor,
										const int IsComplete,
										const int nCounter)
{

   double dHltEstimated;
   double dCommDisplacement = 0.0;
   double dAlphaZero = 0.0;
   double dBetaSorpt = 0.0;
   double dAlphaSorpt = 0.0;
   double dAlphaComb = 0.0;
   double dInitConc = 0.0;
   double dBetaSorptPrime = 0.0;
   int nAdjComputed;
   
   
  
   if ( nCounterPrev != nCounter )
   {
		nAdjComputed = 0;
   }

    //PAT 1/26/06 - If the criteria is complete and CT was not defined by the USER
	if ( dUserDefined == FALSE ) 
	{
	   if ( IsComplete == 1 ) 
	   {
		   if ( nAdjComputed == 0 ) //PAT 1/26/06 - check if the Adj HLT due to Sorption has been computed
		   {
				nCounterPrev = nCounter;
				//dCommDisplacement = GetCommodityDisplacement (dVolume, dLoadFactor, dPorosity); 
				dAlphaZero = GetAlphaBasedonHLT0 ( dHltOriginal );
				dBetaSorpt = GetBetaSorption ( dTemp, dBConstant, dCConstant, dDCZero );
				dAlphaSorpt = GetAlphaSorption ( dLoadFactor, dBetaSorpt );
				dAlphaComb = dAlphaZero + dAlphaSorpt;
				dAdjHLT = GetAlphaBasedonHLT0 ( dAlphaComb );
		   }
	   }
	   else
	   {
			dHltEstimated = GetEstimatedHlt( dHltOriginal );
	   }
	}

   //12/23/04-using the flag, if user has defined a target ct, call GetInitialConcetration2
   //passing orignal HLT; else use GetInitialConcentration
   if ( dUserDefined )
   {
		//Pat 5/5/05 - Nvalue should always be 1 if dUserDefined is true
		m_pestCurrent.dNValue = 1.0;
		return( ::GetInitialConcentration2( dHltOriginal, dTimeExposure, dFinalDose, m_pestCurrent.dNValue ) );
   }
   else
   {
	   if ( IsComplete == 1 )
	   {
		   if ( nAdjComputed == 0 )
		   {
				dAlphaSorpt = 0.0;
				dAlphaComb = 0.0;
				dInitConc = (::GetInitialConcentration( dAdjHLT, dTimeExposure, dFinalDose, m_pestCurrent.dNValue ) );
				dBetaSorptPrime = GetBetaSorptionPrime ( dTemp, dBConstant, dCConstant, dDConstant, dInitConc ); 
				dAlphaSorpt = GetAlphaSorption ( dLoadFactor, dBetaSorptPrime );
				dAlphaComb = dAlphaZero + dAlphaSorpt;
				dAdjHLT = GetAlphaBasedonHLT0 ( dAlphaComb );
				nAdjComputed = 1; //Reset value to 1 so that it would not compute again for the Adj HLT due to Sorption
		   }
			
		   return (::GetInitialConcentration( dAdjHLT, dTimeExposure, dFinalDose, m_pestCurrent.dNValue ) );
	   }
	   else
	   {
			return( ::GetInitialConcentration( dHltEstimated, dTimeExposure, dFinalDose, m_pestCurrent.dNValue ) );
	   }
   }
}


double DDoseCalc::GetConcentrationTime( const double &dHltOriginal, 
                                        const double &dTimeExposure, 
                                        const double &dInitialConcentration )
{
   // Equation for Target Concentration Time (CT), per Part I Step 9, on 2002-07-10
   //
   //       HLT * Co           -(t / HLT)
   // CT  = -------- * [ 1 - 2            ]
   //        ln(2)

   double d2RaisedPower = pow( 2, - ( dTimeExposure / dHltOriginal ) );
   
   return( ( ( dHltOriginal * dInitialConcentration ) / log( 2 ) ) * ( 1 - d2RaisedPower ) );
}


double DDoseCalc::GetMaxInitConcentration( const double &dMaxConcentrationTime,
                                           const double &dHltOriginal, 
                                           const double &dTimeExposure )
{
   // Equation for Maximum Permitted Initial Concentration, per Part I Step 10B, on 2002-07-10
   //
   //              MaxPermittedCT
   // Co  = ----------------------------
   //       HLT            (-t / HLT)
   //       ----- * [ 1 - 2          ]
   //       ln(2)
   
   
   double d2RaisedPower = pow( 2, - dTimeExposure / dHltOriginal );
   
   return( dMaxConcentrationTime / ( ( ( dHltOriginal / log( 2 ) ) * ( 1 - d2RaisedPower ) ) ) );
}


double DDoseCalc::GetMaxDose( const double &dHltOriginal, 
                              const double &dTimeExposure, 
                              const double &dInitialConcentration,
							  const int IsComplete,
							  bool dUserDefined)
{
   // The following is the function to determine the concentration time, per Part I Step 10C, on 2002-07-10
   //                   n
   //          HLTe * Co           -(t / HLTe)
   // Dose  = ---------- * [ 1 - 2            ]
   //            ln(2)

   double dHltEstimated;

   //PAT 1/26/06 - If the criteria is complete and CT was not defined by the USER
   if ( ( IsComplete == 1 ) && ( dUserDefined == FALSE ) )
   {
	   dHltEstimated = dAdjHLT;
   }
   else
   {
	   dHltEstimated = GetEstimatedHlt( dHltOriginal );
   }

   double d2RaisedPower = pow( 2, - ( dTimeExposure / dHltEstimated ) );

   double dInitConcRaisedPower = pow( dInitialConcentration, m_pestCurrent.dNValue );
   
   return( ( ( dHltEstimated * dInitConcRaisedPower ) / log( 2 ) ) * ( 1 - d2RaisedPower ) );
}


//Pat 1/17/06 - Added these unctions that would be needed to compute for the Adj HLT
double DDoseCalc::GetCommodityDisplacement ( const double &dVolume, 
											 const double &dLoadFactor, 
									         const double &dPorosity)
{
	double dCommodityDisplacement = dVolume * (  100 - dLoadFactor  + dPorosity * dLoadFactor ) / 100;
	return ( dCommodityDisplacement );
}

double DDoseCalc::GetAlphaBasedonHLT0( const double &dHltOriginal )
{
	return ( log( 2 ) / dHltOriginal );
}

double DDoseCalc::GetBetaSorption( const double &dTemp,
								   const double &dBConstant,
								   const double &dCConstant,
								   const double &dDCZero )
{
	return  - 1 * ( dBConstant + dCConstant * dTemp + dDCZero );
}

double DDoseCalc::GetAlphaSorption( const double &dLoadFactor,
								    const double dBetaSorpt)
{
	return ( ( dBetaSorpt / 0.8 ) * ( dLoadFactor / 100 ) );
}

double DDoseCalc::GetBetaSorptionPrime ( const double &dTemp,
										 const double &dBConstant,
										 const double &dCConstant,
										 const double &dDConstant,
										 const double dCnTTarget)
{
	return - 1 * ( dBConstant + dCConstant * dTemp + dDConstant * dCnTTarget);
}
//Pat 1/17/06 End of Changes