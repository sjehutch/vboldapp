// DasPestData.cpp: implementation of the DDasPestData class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "DasPestData.h"
#include "GlobalEnums.h"


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DDasPestData::DDasPestData()
{
   // Ensure it is empty to start with
   m_aPestData.RemoveAll();
}


DDasPestData::~DDasPestData()
{
   // Clean up the array
   m_aPestData.RemoveAll();
}


void DDasPestData::PopulatePestData( const PEST &ePest, 
                                     const LIFE_STAGE &eLifeStage, 
                                     const TREATMENT_TYPE &eTreatmentType,
                                     const FUMIGATION_TYPE &eFumigationType,
                                     const double &dTimeExposure,
                                     CString &strMsg )
{
   // Clear out the existing data
   m_aPestData.RemoveAll();

   // Populate the pest array with the selected pest's data
   switch( ePest )
   {
   case PEST_CONFUSED_FLOUR_BEETLE:
      PopulateConfusedFlourBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_RED_FLOUR_BEETLE:
      PopulateRedFlourBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_SAWTOOTHED_GRAIN_BEETLE:
      PopulateSawToothedGrainBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_WAREHOUSE_BEETLE:
      PopulateWarehouseBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_OTHER_BEETLE:
	  //JNM 04/16/03 - Get the parameters of the most tolerant pest
	  if(eLifeStage == EGG_ALL)
		  PopulateRedFlourBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      else
		  PopulateWarehouseBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );	
	  break;

   case PEST_INDIAN_MEAL_MOTH:
      PopulateIndianMealMoth( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_MEDITERRANEAN_FLOUR_MOTH:
      PopulateMedFlourMoth( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_CODLING_MOTH:
      PopulateCodlingMoth( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_NAVEL_ORANGEWORM:
      PopulateNavelOrangeWorm( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_OTHER_MOTH:
      PopulateMedFlourMoth( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   //JNM 04/01/03 - Fum. Enhancements 2003 - Added new case for new pest.
   case PEST_GRANARY_WEEVIL:
      PopulateGranaryWeevil( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   //JNM 04/01/03 - Fum. Enhancements 2003 - Added new case for new pest.
   case PEST_RICE_WEEVIL:
      PopulateRiceWeevil( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   //Pat 4/6/05: Added new case for new pest
   case	PEST_LESSER_GRAIN_BORER:
	  PopulateLesserGrainBorer( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   //Pat 12/2/05 : Enhancement 2005-2006 - Added new pests (5)
   case PEST_HIDE_BEETLE:
	  PopulateHideBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_DRUG_STORE_BEETLE:
	  PopulateDrugStoreBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_RUSTY_GRAIN_BEETLE:
	  PopulateRustyGrainBeetle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_ALMOND_MOTH:
	  PopulateAlmondMoth( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

   case PEST_RODENTS:
	  PopulateRodents( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;


	case PEST_CIGARETTE_BEETLE:
	  PopulateCigarette_Bettle( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

	case PEST_COWPEA_WEEVIL:
	  PopulateCowpeaWeevil( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;


	case PEST_BEAN_WEEVIL:
	  PopulateBeanWeevil( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;

  
	case PEST_COCOA_MOTH:
	  PopulateCocoaMoth( eLifeStage, eTreatmentType, eFumigationType, dTimeExposure, strMsg );
      break;




   default:
      // We should not be here
      ASSERT( FALSE );
   }
}


PEST_DATA DDasPestData::GetPestData( const double &dTemp, BOOL bLowerValue )
{
   PEST_DATA pest;
   
   int nTotal = m_aPestData.GetSize();

   for( int nIndex = 0; nIndex < nTotal; nIndex++ )
   {
      pest = m_aPestData.GetAt( nIndex );

      // Check if the temperature is greater than what is set by user
      // If it is we have found the correct data set and should end the search
      if( pest.dTemp > dTemp )
      {
         // We are trying to get the lower value so we may have to back up 1
         if( bLowerValue )
         {
            // If the index is greater than 0 get the previous data set
            // otherwise we are as low as we can go
            if( nIndex > 0 )
            {
               pest = m_aPestData.GetAt( nIndex - 1 );
            }

            break;
         }
         else
         {
            // We have found the correct pest data
            break;
         }
      }
   }

   return pest;
}


// ****************************************************************************
//
// **********   Our pest data   **********
//
// ****************************************************************************
void DDasPestData::PopulateRedFlourBeetle( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 0.873;
         pest.dIntercept = -14.0253;
         pest.dSlope     = 2.2303;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_SPACE;
				//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
				pest.dSafetyFactor = 1.0;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 0.873;
         pest.dIntercept = -13.528;
         pest.dSlope     = 2.2303;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.0;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 0.746;
         pest.dIntercept = -11.312;
         pest.dSlope     = 2.2973;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );





		 //Pat 12/1/2005 Enhancement 2005-2006 - Expanded temperature 35 and 40
		 pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.41;
         pest.dSlope     = 1.51;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
            pest.dSafetyFactor = 1.0;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = 1.0;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 
		 pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.96;
         pest.dSlope     = 1.61;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				pest.dSafetyFactor = 1.5;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 1.1;
			 }
		 else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );
		 // ---- End of Enhancement 2005-2006 ----


      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.16;
         pest.dSlope     = 2.3;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
             //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.0;
			 pest.dSafetyFactor = 3.5;
         }
         else
         {
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
            //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;
			  pest.dSafetyFactor = 3.5;
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.99;
         pest.dSlope     = 2.72;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
            //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			//CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			//pest.dSafetyFactor = 3.0;
			 pest.dSafetyFactor = 5.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.5;
			 pest.dSafetyFactor = 5.0;
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.95;
         pest.dSlope     = 2.04;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 
			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 5.0;
         }
         else
         {
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
            //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			//pest.dSafetyFactor = 3.5;
			  pest.dSafetyFactor = 5.0;
         }

         m_aPestData.Add( pest );



		 // 03/14/2007 David Smith Added pest data for 35 Degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.92;
         pest.dSlope     = 1.7;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
		    //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
		      pest.dSafetyFactor = 4.15;
         }
         else
         {
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
            //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			//pest.dSafetyFactor = 3.5;
			  pest.dSafetyFactor = 4.15;
         }

         m_aPestData.Add( pest );


		 // 03/14/2007 David Smith Added pest data for 40 Degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.67;
         pest.dSlope     = 1.75;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 
			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 4.2;
         }
         else
         {
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
            //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			//pest.dSafetyFactor = 3.5;
			  pest.dSafetyFactor = 4.2;
         }

         m_aPestData.Add( pest );






      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


void DDasPestData::PopulateSawToothedGrainBeetle( const LIFE_STAGE &eLifeStage, 
                                                  const TREATMENT_TYPE &eTreatmentType, 
                                                  const FUMIGATION_TYPE &eFumigationType,
                                                  const double &dTimeExposure,
                                                  CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 0.873;
         pest.dIntercept = -14.0253;
         pest.dSlope     = 2.2303;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/30/03 - PT fix - STGB should have same value as RFB
            pest.dSafetyFactor = 1.0;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 0.873;
         pest.dIntercept = -13.528;
         pest.dSlope     = 2.2303;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/30/03 - PT fix - STGB should have same value as RFB
            pest.dSafetyFactor = 1.0;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 0.746;
         pest.dIntercept = -11.312;
         pest.dSlope     = 2.2973;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/30/03 - PT fix - STGB should have same value as RFB
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );




		 // 03-15-07 David Smith Adding data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1;
         pest.dIntercept = -8.41;
         pest.dSlope     = 1.51;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/30/03 - PT fix - STGB should have same value as RFB
            pest.dSafetyFactor = 1.0;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = 1.0;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 // 03-15-07 David Smith Adding data for 40 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1;
         pest.dIntercept = -7.96;
         pest.dSlope     = 1.61;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/30/03 - PT fix - STGB should have same value as RFB
            pest.dSafetyFactor = 1.5;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = 1.1;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );





















      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.36;
         pest.dSlope     = 1.57;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 2.0;
			   pest.dSafetyFactor = 6.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 6.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.01;
         pest.dSlope     = 2.54;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 
			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 2.0;
			   pest.dSafetyFactor = 6.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 6.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.52;
         pest.dSlope     = 1.94;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			  //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			  //pest.dSafetyFactor = 2.0;
			    pest.dSafetyFactor = 6.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 6.0;
         }

         m_aPestData.Add( pest );



		 // 03-15-07 David Smith Adding data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1;
         pest.dIntercept = -4.52;
         pest.dSlope     = 1.94;
         pest.dProbit    = PROBIT_95;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
           		pest.dSafetyFactor = 5.0;
			 }
         else
			 if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
				 }
			 else
				 {
					ASSERT( FALSE );
				 }

         m_aPestData.Add( pest );


		 // 03-15-07 David Smith Adding data for 35 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1;
         pest.dIntercept = -4.52;
         pest.dSlope     = 1.94;
         pest.dProbit    = PROBIT_95;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
           		pest.dSafetyFactor = 4.2;
			 }
         else
			 if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
				 }
			 else
				 {
					ASSERT( FALSE );
				 }

         m_aPestData.Add( pest );














      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


void DDasPestData::PopulateMedFlourMoth( const LIFE_STAGE &eLifeStage, 
                                         const TREATMENT_TYPE &eTreatmentType, 
                                         const FUMIGATION_TYPE &eFumigationType,
                                         const double &dTimeExposure,
                                         CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
		 //JNM 04/20/03 - Removed data for temp 15 per new instructions
         //pest.dTemp      = 15;
         //pest.dNValue    = 1.3;
         //pest.dIntercept = -20.034;
         //pest.dSlope     = 2.4637;
         //if( eFumigationType == SPACE_FUMIGATION )
         //{
         //   pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
         //   pest.dSafetyFactor = 1.0;
         //}
         //else
         //if( eFumigationType == COMMODITY_FUMIGATION )
         //{
         //   pest.dProbit    = PROBIT_95;
         //   pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         //}
         //else
         //{
         //   ASSERT( FALSE );
         //}

         //m_aPestData.Add( pest );

   
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -24.209;
         pest.dSlope     = 3.5886;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.1;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.245;
         pest.dIntercept = -13.395;
         pest.dSlope     = 2.1153;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//JNM 04/20/2003 - Changed safety factor value again per new
			//		instructions.
            pest.dSafetyFactor = 1.5;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );

   
         
		 
		 
		 pest.dTemp      = 30;
         pest.dNValue    = 1.271;
         pest.dIntercept = -15.515;
         pest.dSlope     = 2.6312;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.5;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


		 // 03/15/07 David Smith  Adding data for 35 degrees
		 pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -15.515;
         pest.dSlope     = 2.6312;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 0.6;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = 0.47;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );




		 // 03/15/07 David Smith  Adding data for 35 degrees
		 pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -15.515;
         pest.dSlope     = 2.6312;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 0.52;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = 0.41;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );




      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.13;
         pest.dSlope     = 2.35;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 4.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 4.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.71;
         pest.dSlope     = 2.16;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			//CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			//pest.dSafetyFactor = 3.0;
			  pest.dSafetyFactor = 5.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 5.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.15;
         pest.dSlope     = 2.47;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			//pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 5.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
		    //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 5.0;
         }

         m_aPestData.Add( pest );




		 // 03/15/07 David Smith Adding Pest data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.15;
         pest.dSlope     = 2.47;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				pest.dSafetyFactor = 3.5;
			 }
         else
			 {
				pest.dSafetyFactor = 5.0;
			 }

         m_aPestData.Add( pest );


		 // 03/15/07 David Smith Adding Pest data for 40 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.15;
         pest.dSlope     = 2.47;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				pest.dSafetyFactor = 3.0;
			 }
         else
			 {
				pest.dSafetyFactor = 5.0;
			 }

         m_aPestData.Add( pest );












      }
   }
   else
   {
      ASSERT( FALSE );
   }
}



void DDasPestData::PopulateIndianMealMoth( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType, 
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 0.5222;
         pest.dIntercept = -4.31897;
         pest.dSlope     = 1.2354;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.8;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 0.5222;
         pest.dIntercept = -4.0433;
         pest.dSlope     = 1.2354;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.8 to 1.9
            pest.dSafetyFactor = 1.9;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 0.5434;
         pest.dIntercept = -5.3295;
         pest.dSlope     = 1.5141;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.5 to 2.0
            pest.dSafetyFactor = 1.7;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );

        
		 // Adding pest data for 35 degrees 
		 pest.dTemp      = 35;
         pest.dNValue    = 1;
         pest.dIntercept = -5.3295;
         pest.dSlope     = 1.5141;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.5 to 2.0
            pest.dSafetyFactor = 6.3;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = 4.2;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 // Adding pest data for 40 degrees 
		 pest.dTemp      = 40;
         pest.dNValue    = 1;
         pest.dIntercept = -5.3295;
         pest.dSlope     = 1.5141;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.5 to 2.0
            pest.dSafetyFactor = 5.5;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = 3.7;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );

















      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.19;
         pest.dSlope     = 1.34;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 4.7;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 4.7;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.33;
         pest.dSlope     = 1.67;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.2;
			   pest.dSafetyFactor = 4.7;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.7;
			   pest.dSafetyFactor = 4.7;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.04;
         pest.dSlope     = 1.94;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 2.3;
			   pest.dSafetyFactor = 3.3;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 2.8;
			   pest.dSafetyFactor = 3.3;
         }

         m_aPestData.Add( pest );






		 // 3-15-2007 David Smith Adding Pest Data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.04;
         pest.dSlope     = 1.94;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            pest.dSafetyFactor = 2.5;
         }
         else
         {
           pest.dSafetyFactor = 3.3;
         }

         m_aPestData.Add( pest );


		 // 3-15-2007 David Smith Adding Pest Data for 40 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.04;
         pest.dSlope     = 1.94;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            pest.dSafetyFactor = 2.0;
         }
         else
         {
           pest.dSafetyFactor = 3.3;
         }

         m_aPestData.Add( pest );
















      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


void DDasPestData::PopulateConfusedFlourBeetle( const LIFE_STAGE &eLifeStage, 
                                                const TREATMENT_TYPE &eTreatmentType, 
                                                const FUMIGATION_TYPE &eFumigationType,
                                                const double &dTimeExposure,
                                                CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.303;
         pest.dIntercept = -13.442;
         pest.dSlope     = 1.9984;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//JNM 04/20/2003 - Changed safety factor value again.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.5 to 1.6
            pest.dSafetyFactor = 1.6;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 //Inserts pest data for temp 25 degrees
         pest.dTemp      = 25;
         pest.dNValue    = 1.303;
         pest.dIntercept = -14.360;
         pest.dSlope     = 2.2292;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//JNM 04/14/2003 - Changed safety factor value again.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.5 to 1.6
            pest.dSafetyFactor = 1.6;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


          //Inserts pest data for temp 25 degrees
         pest.dTemp      = 30;
         pest.dNValue    = 1.303;
         pest.dIntercept = -18.837;
         pest.dSlope     = 2.8365;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//JNM 04/14/2003 - Changed safety factor value again.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.4 to 1.5
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );






		//David Smith 3-15-07 Added new data for temp 35
          //Inserts pest data for temp 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1;
         pest.dIntercept = -11.1;
         pest.dSlope     = 2.1;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
		    pest.dSafetyFactor = 1.1;
         }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 1.1;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );



		//David Smith 3-15-07 Added new data for temp 35
          //Inserts pest data for temp 40 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1;
         pest.dIntercept = -11.1;
         pest.dSlope     = 2.1;
         if( eFumigationType == SPACE_FUMIGATION )
				 {
					pest.dProbit    = PROBIT_85;
					pest.dSafetyFactor = 1.0;
				 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dProbit    = PROBIT_95;
					pest.dSafetyFactor = 0.7;
				 }
		 else
				 {
					ASSERT( FALSE );
				 }

         m_aPestData.Add( pest );
















      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.21;
         pest.dSlope     = 1.82;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 3.3;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 3.3;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.96;
         pest.dSlope     = 1.79;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 pest.dSafetyFactor = 3.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 3.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.47;
         pest.dSlope     = 1.39;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 4.0;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 4.0;
         }

         m_aPestData.Add( pest );



		 //03/14/07 David Smith Adding pest data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.92;
         pest.dSlope     = 1.7;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 4.2;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 4.0;
         }

         m_aPestData.Add( pest );




		 //03/14/07 David Smith Adding pest data for 35 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.67;
         pest.dSlope     = 1.75;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 3.0;
			   pest.dSafetyFactor = 4.2;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 3.5;
			   pest.dSafetyFactor = 4.0;
         }

         m_aPestData.Add( pest );







      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


void DDasPestData::PopulateWarehouseBeetle( const LIFE_STAGE &eLifeStage, 
                                            const TREATMENT_TYPE &eTreatmentType, 
                                            const FUMIGATION_TYPE &eFumigationType,
                                            const double &dTimeExposure,
                                            CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.177;
         pest.dIntercept = -20.517;
         pest.dSlope     = 2.8017;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 0.7385;
         pest.dIntercept = -10.934;
         pest.dSlope     = 2.1672;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.41;
         pest.dSlope     = 1.23;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
			//CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			// from 1.2 to 1.4
            pest.dSafetyFactor = 1.4;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


		 //03/15/2007 David Smith Adding data for temp 35
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.41;
         pest.dSlope     = 1.51;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				pest.dSafetyFactor = 1.0;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 1.0;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );



		 //03/15/2007 David Smith Adding data for temp 35
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.96;
         pest.dSlope     = 1.61;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				pest.dSafetyFactor = 1.5;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 1.1;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );




      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.87;
         pest.dSlope     = 2.28;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.0;
			   pest.dSafetyFactor = 2.8;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.8;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.87;
         pest.dSlope     = 2.28;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.0;
			   pest.dSafetyFactor = 2.44;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.8;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.71;
         pest.dSlope     = 2.05;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.0;
			   pest.dSafetyFactor = 2.05;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			 pest.dSafetyFactor = 2.8;
         }

         m_aPestData.Add( pest );





		 //03/15/2007 David Smith Adding data for temp 35
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.92;
         pest.dSlope     = 1.7;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			pest.dSafetyFactor = 4.15;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



 //03/15/2007 David Smith Adding data for temp 40
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.67;
         pest.dSlope     = 1.75;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			pest.dSafetyFactor = 4.25;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );








      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


// JNM 04/01/03 - Fumiguide Enhancements 2003 - Added this new method for the new pest,
//			Granary Weevil. This is the exact copy of the existing method 
//			PopulateWarehouseBeetle except that the values for the following 
//			were changed: dTemp, dNValue, dIntercept, dSlope.
void DDasPestData::PopulateGranaryWeevil( const LIFE_STAGE &eLifeStage, 
                                            const TREATMENT_TYPE &eTreatmentType, 
                                            const FUMIGATION_TYPE &eFumigationType,
                                            const double &dTimeExposure,
                                            CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.16;
         pest.dIntercept = -12.3;
         pest.dSlope     = 1.84;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.11;
         pest.dIntercept = -12.8;
         pest.dSlope     = 2.07;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 1.15;
         pest.dIntercept = -10.2;
         pest.dSlope     = 1.77;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 // 3-15-07 David Smith Added 35  and 35 degree tmperature
         pest.dTemp      = 35;
         pest.dNValue    = 1;
         pest.dIntercept = -10.2;
         pest.dSlope     = 1.77;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				pest.dSafetyFactor = 0.7;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 0.6;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );


         pest.dTemp      = 40;
         pest.dNValue    = 1;
         pest.dIntercept = -10.2;
         pest.dSlope     = 1.77;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				pest.dSafetyFactor = 0.62;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 0.52;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );






















      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -0.83;
         pest.dSlope     = 0.656;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -3.68;
         pest.dSlope     = 1.99;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -3.68;
         pest.dSlope     = 1.99;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;
			 
			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.0;
         }

         m_aPestData.Add( pest );


		 // 03/15/07 David Smith	Adding data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -3.68;
         pest.dSlope     = 1.99;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				 pest.dSafetyFactor = 2.0;
			 }
         else
			 {
				pest.dSafetyFactor = 2.0;
			 }

         m_aPestData.Add( pest );



		 // 03/15/07 David Smith	Adding data for 40 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -3.68;
         pest.dSlope     = 1.99;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				 pest.dSafetyFactor = 2.0;
			 }
         else
			 {
				pest.dSafetyFactor = 2.0;
			 }

         m_aPestData.Add( pest );


















      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


// JNM 04/01/03 - Fumiguide Enhancements 2003 - Added this new method for the new pest,
//			Rice Weevil. This is the exact copy of the existing method 
//			PopulateWarehouseBeetle except that the values for the following 
//			were changed: dTemp, dNValue, dIntercept, dSlope.
void DDasPestData::PopulateRiceWeevil( const LIFE_STAGE &eLifeStage, 
                                            const TREATMENT_TYPE &eTreatmentType, 
                                            const FUMIGATION_TYPE &eFumigationType,
                                            const double &dTimeExposure,
                                            CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 0.942;
		 //JNM 04/20/2003 - Corrected intercept value.
         pest.dIntercept = -12.9;
         pest.dSlope     = 2.06;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.0;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 0.869;
         pest.dIntercept = -13.5;
         pest.dSlope     = 2.39;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.0;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 0.965;
         pest.dIntercept = -10.4;
         pest.dSlope     = 1.9;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85;
			//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
            pest.dSafetyFactor = 1.2;
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95;
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 //03-15-07 David Smith Added data fro 35 and 40 tempature
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.12;
         pest.dSlope     = 1.05;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				pest.dSafetyFactor = 1.2;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 1.0;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );


		 pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -10.4;
         pest.dSlope     = 1.9;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85;
				//JNM 04/14/2003 - Fum. Enh. 2003 - Change safety factor value.
				pest.dSafetyFactor = 0.9;
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95;
				pest.dSafetyFactor = 0.75;
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );



























      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.6;
         pest.dSlope     = 1.41;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.6;
         pest.dSlope     = 1.58;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.0;
         }

         m_aPestData.Add( pest );

   
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.6;
         pest.dSlope     = 1.58;
         pest.dProbit    = PROBIT_95;
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
            //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE;
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
             //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
			 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE;

			 //CG 03/03/2004 - Ticket # 2120829:Change safety factor value 
			 //pest.dSafetyFactor = 2.5;
			   pest.dSafetyFactor = 2.0;
         }

         m_aPestData.Add( pest );

		 // 03/15/2007 David Smith Adding pest data for 35 degree
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.6;
         pest.dSlope     = 1.58;
         pest.dProbit    = PROBIT_95;
         pest.dSafetyFactor = 1.72;
		
		  m_aPestData.Add( pest );


		 // 03/15/2007 David Smith Adding pest data for 40 degree
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.6;
         pest.dSlope     = 1.58;
         pest.dProbit    = PROBIT_95;
         pest.dSafetyFactor = 1.35;
		
		 m_aPestData.Add( pest );







      }
   }
   else
   {
      ASSERT( FALSE );
   }
}



void DDasPestData::PopulateCodlingMoth( const LIFE_STAGE &eLifeStage, 
                                        const TREATMENT_TYPE &eTreatmentType, 
                                        const FUMIGATION_TYPE &eFumigationType,
                                        const double &dTimeExposure,
                                        CString &strMsg )
{
   ASSERT( eFumigationType ); // Removing debug build warning
   ASSERT( dTimeExposure );   // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL || eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.95;
         pest.dSlope     = 1.72;
         pest.dProbit    = PROBIT_95;
         
		 //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
		 //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC;
		 pest.dSafetyFactor = 1.3;

         m_aPestData.Add( pest );
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }

   if( eLifeStage == EGG_ALL )
   {
      // Need to add a warning message
      strMsg.LoadString( IDS_WARNING_CM_EGGS );
   }
}


void DDasPestData::PopulateNavelOrangeWorm( const LIFE_STAGE &eLifeStage, 
                                            const TREATMENT_TYPE &eTreatmentType, 
                                            const FUMIGATION_TYPE &eFumigationType,
                                            const double &dTimeExposure,
                                            CString &strMsg )
{
   ASSERT( eFumigationType ); // Removing debug build warning
   ASSERT( dTimeExposure );   // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest

   if( eLifeStage == EGG_ALL || eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -2.1;
         pest.dSlope     = 0.747;
         pest.dProbit    = PROBIT_95;

		 //CG 01/27/2004 - Launch Changes 2004 - Change safety factor value
         //pest.dSafetyFactor = DOSE_SAFETY_FACTOR_POST_EMBRYONIC;
		 pest.dSafetyFactor = 1.3;

         m_aPestData.Add( pest );
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }


   if( eLifeStage == EGG_ALL )
   {
      // Need to add a warning message
      strMsg.LoadString( IDS_WARNING_NOW_EGGS );
   }
}

//Pat 4/6/05 : Added new pest
void DDasPestData::PopulateLesserGrainBorer( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  // Removing debug build warning

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.38;
         pest.dSlope     = 1.64;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.3; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.06;
         pest.dSlope     = 1.56;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.5; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.27;
         pest.dSlope     = 1.62;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 2.5; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = DOSE_SAFETY_FACTOR_EGGS_COMMODITY; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 15;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.07;
         pest.dSlope     = 1.87;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 3.0;
         }
         else
         {
			  pest.dSafetyFactor = 3.0; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.59;
         pest.dSlope     = 1.81;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 2.5; 
         }
         else
         {
			 pest.dSafetyFactor = 2.5; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.11;
         pest.dSlope     = 1.81;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			   pest.dSafetyFactor = 3.0; 
         }
         else
         {
			  pest.dSafetyFactor = 3.0; 
         }

         m_aPestData.Add( pest );
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


//Pat 12/1/2005: Enhancement 2005-2006 Added new Pest
void DDasPestData::PopulateHideBeetle( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg ); 

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.2;
         pest.dSlope     = 1.52;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.2; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.1; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         // HideBeetle data for 25 degrees
		 pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -10.1;
         pest.dSlope     = 1.67;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.2; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.1; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


		 //03-15-07 Daivd Smith Modified safty factor per instrucntion from Suresh
		 pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.95;
         pest.dSlope     = 1.38;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.5; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.2; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 //03-15-07 Daivd Smith Modified safty factor per instrucntion from Suresh
		 pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.95;
         pest.dSlope     = 1.38;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.84; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );




		 		 //03-15-07 Daivd Smith Modified safty factor per instrucntion fro Suresh
		 pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.95;
         pest.dSlope     = 1.38;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 0.85; 
			 }
		 else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95; 
				pest.dSafetyFactor = 0.6; 
			 }
		 else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );







      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -10.2;
         pest.dSlope     = 2.46;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 2.5;
         }
         else
         {
			  pest.dSafetyFactor = 2.5; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.98;
         pest.dSlope     = 2.03;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 2.5; 
         }
         else
         {
			 pest.dSafetyFactor = 2.5; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.67;
         pest.dSlope     = 1.96;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			   pest.dSafetyFactor = 2.5; 
         }
         else
         {
			  pest.dSafetyFactor = 2.5; 
         }

         m_aPestData.Add( pest );


		 // 03/15/2007 David Smith Added pest data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.67;
         pest.dSlope     = 1.96;
         pest.dProbit    = PROBIT_95; 
         pest.dSafetyFactor = 2.15; 
         
		 m_aPestData.Add( pest );

		 // 03/15/2007 David Smith Added pest data for 40 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.67;
         pest.dSlope     = 1.96;
         pest.dProbit    = PROBIT_95; 
         pest.dSafetyFactor = 1.8; 
         
		 m_aPestData.Add( pest );











      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}

//Pat 12/1/2005: Enhancement 2005-2006 Added new Pest
void DDasPestData::PopulateDrugStoreBeetle( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg ); 

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.88;
         pest.dSlope     = 1.6;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.2; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.59;
         pest.dSlope     = 1.63;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.3; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.1; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );

		 // DrugStoreBettle data for 30 degrees
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.17;
         pest.dSlope     = 1.26;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.5; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.5; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


		 //Adding tempature 35  
		 pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.17;
         pest.dSlope     = 1.26;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.93; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 
		 //Adding tempature 40  
		 pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.17;
         pest.dSlope     = 1.26;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 0.9; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.84; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );









      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.2;
         pest.dSlope     = 2.25;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 2.5;
         }
         else
         {
			  pest.dSafetyFactor = 2.5; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.82;
         pest.dSlope     = 2.27;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			pest.dSafetyFactor = 2.85; 
         }
         else
         {
			pest.dSafetyFactor = 2.85;
         }

         m_aPestData.Add( pest );


		 // DrugStoreBettle data for 30 degrees
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.24;
         pest.dSlope     = 1.79;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			pest.dSafetyFactor = 2.25; 
         }
         else
         {
			pest.dSafetyFactor = 2.25; 
         }

         m_aPestData.Add( pest );


		 // 03/15/07 David Smith adding DrugStoreBettle data for 35 degrees
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.24;
         pest.dSlope     = 1.79;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			pest.dSafetyFactor = 1.95; 
         }
         else
         {
			pest.dSafetyFactor = 1.95; 
         }

         m_aPestData.Add( pest );



		 // 03/15/07 David Smith adding DrugStoreBettle data for 40 degrees
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.24;
         pest.dSlope     = 1.79;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			pest.dSafetyFactor = 1.7; 
         }
         else
         {
			pest.dSafetyFactor = 1.7; 
         }

         m_aPestData.Add( pest );

      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}

//Pat 12/2/2005: Enhancement 2005-2006 Added new Pest
void DDasPestData::PopulateRustyGrainBeetle( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg ); 

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 15;
         pest.dNValue    = 1.0;
         pest.dIntercept = -13.3;
         pest.dSlope     = 2.04;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 1.1; 
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95; 
				pest.dSafetyFactor = 0.95; 
			 }
         else
			 {
				ASSERT( FALSE );
			 }

         m_aPestData.Add( pest );


         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.01;
         pest.dSlope     = 1.44;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.1; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.9; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.47;
         pest.dSlope     = 1.64;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.3; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.0;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



		 // Addint data for temps 30 35 40 David Smith 03/15/07
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.47;
         pest.dSlope     = 1.64;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 0.76; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.7;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.47;
         pest.dSlope     = 1.64;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 0.63; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.48;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );



         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -9.47;
         pest.dSlope     = 1.64;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 0.54; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.35;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 15;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.13;
         pest.dSlope     = 1.59;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				 pest.dSafetyFactor = 3.0;
			 }
         else
			 {
				 pest.dSafetyFactor = 3.0; 
			 }

         m_aPestData.Add( pest );


		 // RustyGrainBeetle data for 20 degrees
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -3.73;
         pest.dSlope     = 1.6;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				 pest.dSafetyFactor = 3.0; 
			 }
         else
			 {
				 pest.dSafetyFactor = 3.0; 
			 }

         m_aPestData.Add( pest );


         // RustyGrainBeetle data for 25 degrees
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.93;
         pest.dSlope     = 2.35;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				pest.dSafetyFactor = 3.0; 
			 }
         else
			 {
				pest.dSafetyFactor = 3.0; 
			 }

         m_aPestData.Add( pest );


         // RustyGrainBeetle data for 30 degrees
		 // Added 03/15/07 David Smith
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.93;
         pest.dSlope     = 2.35;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				pest.dSafetyFactor = 2.8; 
			 }
         else
			 {
				pest.dSafetyFactor = 2.8; 
			 }

         m_aPestData.Add( pest );


         // RustyGrainBeetle data for 35 degrees
		 // Added 03/15/07 David Smith
         pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.93;
         pest.dSlope     = 2.35;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				pest.dSafetyFactor = 2.45; 
			 }
         else
			 {
				pest.dSafetyFactor = 2.45; 
			 }

         m_aPestData.Add( pest );



		  // RustyGrainBeetle data for 340 degrees
		 // Added 03/15/07 David Smith
         pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.93;
         pest.dSlope     = 2.35;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
			 {
				pest.dSafetyFactor = 2.2; 
			 }
         else
			 {
				pest.dSafetyFactor = 3.2; 
			 }

         m_aPestData.Add( pest );

      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}

//Pat 12/2/2005: Enhancement 2005-2006 Added new Pest
void DDasPestData::PopulateAlmondMoth( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg ); 

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      =20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -13.1;
         pest.dSlope     = 2.01;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.1; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.97; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -12.5;
         pest.dSlope     = 1.94;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.1; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -7.76;
         pest.dSlope     = 1.46;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.4; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.2;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


		 // 03/17/2007 David Smith Adding pest data for 35 degrees
		 pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -6.28;
         pest.dSlope     = 1.32;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.4; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.2;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


		 // 03/17/2007 David Smith Adding pest data for 40 degrees
		 pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.9;
         pest.dSlope     = 1.93;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
			pest.dSafetyFactor = 1.4; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.4;
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );











      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = -5.31;
         pest.dSlope     = 1.64;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 5.0;
         }
         else
         {
			 pest.dSafetyFactor = 5.0; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = -3.86;
         pest.dSlope     = 1.28;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 4.0; 
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = -4.65;
         pest.dSlope     = 1.63;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			pest.dSafetyFactor = 4.0; 
         }
         else
         {
			pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );



	     // 03/17/2007 David Smith Adding pest data for 35 degrees
		 pest.dTemp      = 35;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.15;
         pest.dSlope     = 2.47;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			pest.dSafetyFactor = 3.5; 
         }
         else
         {
			pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );


	     // 03/17/2007 David Smith Adding pest data for 40 degrees
		 pest.dTemp      = 40;
         pest.dNValue    = 1.0;
         pest.dIntercept = -8.15;
         pest.dSlope     = 2.47;
         pest.dProbit    = PROBIT_95; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			pest.dSafetyFactor = 3.0; 
         }
         else
         {
			pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );






      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}


//Pat 12/1/2005: Enhancement 2005-2006 Added new Pest
void DDasPestData::PopulateRodents ( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   if( eLifeStage == EGG_ALL )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 10;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
			pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 15;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
			pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );
	 
         pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
			pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );
	 
         pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
			pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_99; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

         m_aPestData.Add( pest );
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         pest.dTemp      = 10;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         pest.dProbit    = PROBIT_99; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 1.0;
         }
         else
         {
			 pest.dSafetyFactor = 1.0; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 15;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         pest.dProbit    = PROBIT_99; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 1.0; 
         }
         else
         {
			 pest.dSafetyFactor = 1.0; 
         }

         m_aPestData.Add( pest );


         pest.dTemp      = 20;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         pest.dProbit    = PROBIT_99; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			  pest.dSafetyFactor = 1.0; 
         }
         else
         {
			  pest.dSafetyFactor = 1.0; 
         }

         m_aPestData.Add( pest );

	 
	     pest.dTemp      = 25;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         pest.dProbit    = PROBIT_99; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			  pest.dSafetyFactor = 1.0; 
         }
         else  
         {
			  pest.dSafetyFactor = 1.0; 
         }

         m_aPestData.Add( pest );

	 
		 pest.dTemp      = 30;
         pest.dNValue    = 1.0;
         pest.dIntercept = 36;
         pest.dSlope     = 0;
         pest.dProbit    = PROBIT_99; 
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			  pest.dSafetyFactor = 1.0; 
         }
         else
         {
			  pest.dSafetyFactor = 1.0; 
         }

         m_aPestData.Add( pest );
	 
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}



//David S 03/13/2007: Enhancement 2005-2007 Added new Pest
void DDasPestData::PopulateCigarette_Bettle( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
  
   
   if( eLifeStage == EGG_ALL )// Dosage = "HIGH"
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Inserts data for pest at 20 degrees
         pest.dTemp      = 20;	// temp
         pest.dNValue    = 1.0; // N
         pest.dIntercept = -8.72;  // Intercept
         pest.dSlope     = 1.3;   // Slop

		 // fumigation Type = Space
         if( eFumigationType == SPACE_FUMIGATION ) 
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 0.81; 
			 }
         else // Fumigation Type = Commodity
			 if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dProbit    = PROBIT_95; 
					pest.dSafetyFactor = 0.5; 
				 }
			 else
				 {
					ASSERT( FALSE );
				 }

         // Adds pest data to array
		 m_aPestData.Add( pest );


		 // Inserts data for pest at 25 degrees
         pest.dTemp      =  25;
         pest.dNValue    =  1.0;
         pest.dIntercept = -14.3;
         pest.dSlope     = 2.18;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.75; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


		 // Inserts data for pest at 30 degrees
         pest.dTemp      =  30;
         pest.dNValue    =  1.0;
         pest.dIntercept = -9.95;
         pest.dSlope     = 1.65;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.7; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );




		 // Inserts data for pest at 35 degrees
         pest.dTemp      =  35;
         pest.dNValue    =  1.0;
         pest.dIntercept = -13.8;
         pest.dSlope     = 2.26;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 0.77; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );



		 // Inserts data for pest at 40 degrees
         pest.dTemp      =  40;
         pest.dNValue    =  1.0;
         pest.dIntercept = -11.3;
         pest.dSlope     = 2.17;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.2; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.2; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


 }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Adds pest Data for 20 degrees 
		 pest.dTemp      = 20;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -4.19;         // Interface
         pest.dSlope     = 1.36;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 4.0;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );


		 // Adds pest Data for 25 degrees 
		 pest.dTemp      = 25;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -5.45;         // Interface
         pest.dSlope     = 1.57;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 3.0;
         }
         else
         {
			 pest.dSafetyFactor = 3.0; 
         }

         m_aPestData.Add( pest );


		 // Adds pest Data for 30 degrees 
		 pest.dTemp      = 30;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -6.15;         // Interface
         pest.dSlope     = 1.92;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 4.0;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }
         m_aPestData.Add( pest );





		 // Adds pest Data for 35 degrees 
		 pest.dTemp      = 35;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -6.5;         // Interface
         pest.dSlope     = 2.12;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 4.0;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );

		 
		 // Adds pest Data for 40 degrees 
		 pest.dTemp      = 40;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -6.5;         // Interface
         pest.dSlope     = 2.12;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 3.0;
         }
         else
         {
			 pest.dSafetyFactor = 3.0; 
         }

         m_aPestData.Add( pest );




	 
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}







//David S 03/13/2007: Enhancement 2005-2007 Added new Pest
void DDasPestData::PopulateCowpeaWeevil( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
  
   
   if( eLifeStage == EGG_ALL )// Dosage = "HIGH"
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Inserts data for pest at 20 degrees
         pest.dTemp      = 20;	     // temp
         pest.dNValue    = 1.0;     // N
         pest.dIntercept = -6.18;  // Intercept
         pest.dSlope     = 1.24;   // Slop

		 // fumigation Type = Space
         if( eFumigationType == SPACE_FUMIGATION ) 
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 1.5; 
			 }
         else // Fumigation Type = Commodity
			 if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dProbit    = PROBIT_95; 
					pest.dSafetyFactor = 1.1; 
				 }
			 else
				 {
					ASSERT( FALSE );
				 }

         // Adds pest data to array
		 m_aPestData.Add( pest );


		 // Inserts data for pest at 25 degrees
         pest.dTemp      =  25;
         pest.dNValue    =  1.0;
         pest.dIntercept = -4.61;
         pest.dSlope     = 1.04;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.5; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.3; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


		 // Inserts data for pest at 30 degrees
         pest.dTemp      =  30;
         pest.dNValue    =  1.0;
         pest.dIntercept = -3.43;
         pest.dSlope     = 0.884;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.5; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.4; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );




		 // Inserts data for pest at 35 degrees
         pest.dTemp      =  35;
         pest.dNValue    =  1.0;
         pest.dIntercept = -3.43;
         pest.dSlope     = 0.884;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.2; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.2; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );



		 // Inserts data for pest at 40 degrees
         pest.dTemp      =  40;
         pest.dNValue    =  1.0;
         pest.dIntercept = -3.43;
         pest.dSlope     = 0.884;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


 }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Adds pest Data for 20 degrees 
		 pest.dTemp      = 20;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -0.834;         // Interface
         pest.dSlope     = 0.84;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 9.95;
         }
         else
         {
			 pest.dSafetyFactor = 9.95; 
         }

         m_aPestData.Add( pest );


		 // Adds pest Data for 25 degrees 
		 pest.dTemp      = 25;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = 1.1;         // Interface
         pest.dSlope     = 0.249;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 14.2;
         }
         else
         {
			 pest.dSafetyFactor = 14.2; 
         }

         m_aPestData.Add( pest );


		 // Adds pest Data for 30 degrees 
		 pest.dTemp      = 30;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = 1.09;         // Interface
         pest.dSlope     = 0.259;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 12;
         }
         else
         {
			 pest.dSafetyFactor = 12; 
         }
         m_aPestData.Add( pest );





		 // Adds pest Data for 35 degrees 
		 pest.dTemp      = 35;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = 1.09;         // Interface
         pest.dSlope     = 0.259;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 10;
         }
         else
         {
			 pest.dSafetyFactor = 10; 
         }

         m_aPestData.Add( pest );

		 
		 // Adds pest Data for 40 degrees 
		 pest.dTemp      = 40;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = 1.09;         // Interface
         pest.dSlope     = 0.259;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 8.0;
         }
         else
         {
			 pest.dSafetyFactor = 8.0; 
         }

         m_aPestData.Add( pest );




	 
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}



//David S 03/13/2007: Enhancement 2005-2007 Added new Pest
void DDasPestData::PopulateBeanWeevil( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
   if( eLifeStage == EGG_ALL )// Dosage = "HIGH"
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Inserts data for pest at 20 degrees
         pest.dTemp      = 20;	    // temp
         pest.dNValue    = 1.0;    // N
         pest.dIntercept = -7.34;  // Intercept
         pest.dSlope     = 1.42;   // Slop

		 // fumigation Type = Space
         if( eFumigationType == SPACE_FUMIGATION ) 
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 1.4; 
			 }
         else // Fumigation Type = Commodity
			 if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dProbit    = PROBIT_95; 
					pest.dSafetyFactor = 1.1; 
				 }
			 else
				 {
					ASSERT( FALSE );
				 }

         // Adds pest data to array
		 m_aPestData.Add( pest );


	 // Inserts data for pest at 25 degrees
         pest.dTemp      =  25;
         pest.dNValue    =  1.0;
         pest.dIntercept = -5.78;
         pest.dSlope     = 1.2;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.4; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.2; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


	 // Inserts data for pest at 30 degrees
         pest.dTemp      =  30;
         pest.dNValue    =  1.0;
         pest.dIntercept = -3.88;
         pest.dSlope     = 0.947;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 1.4; 
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95; 
				pest.dSafetyFactor = 1.3; 
			 }
         else
			 {
				ASSERT( FALSE );
			 }

		m_aPestData.Add( pest );




	 //Inserts data for pest at 35 degrees
         pest.dTemp      =  35;
         pest.dNValue    =  1.0;
         pest.dIntercept = -3.88;
         pest.dSlope     = 0.947;
         if( eFumigationType == SPACE_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 1.2; 
			 }
         else if( eFumigationType == COMMODITY_FUMIGATION )
			 {
				pest.dProbit    = PROBIT_95; 
				pest.dSafetyFactor = 1.2; 
			 }
         else
			 {
				ASSERT( FALSE );
			 }

		m_aPestData.Add( pest );



		 // Inserts data for pest at 40 degrees
         pest.dTemp      =  40;
         pest.dNValue    =  1.0;
         pest.dIntercept = -3.88;
         pest.dSlope     = 0.947;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


 }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Adds pest Data for 20 degrees 
	 pest.dTemp      = 20;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -3.93;         // Interface
         pest.dSlope     = 1.35;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 6.0;
         }
         else
         {
			 pest.dSafetyFactor = 6.0; 
         }

         m_aPestData.Add( pest );


	 // Adds pest Data for 25 degrees 
	 pest.dTemp      = 25;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -3.97;         // Interface
         pest.dSlope     = 1.36;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 5.0;
         }
         else
         {
			 pest.dSafetyFactor = 5.0; 
         }

         m_aPestData.Add( pest );


	 // Adds pest Data for 30 degrees 
	 pest.dTemp      = 30;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -2.88;         // Interface
         pest.dSlope     = 1.22;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
		pest.dSafetyFactor = 6.0;
         }
         else
         {
		pest.dSafetyFactor = 6.0; 
         }
         m_aPestData.Add( pest );





	 // Adds pest Data for 35 degrees 
	 pest.dTemp      = 35;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -2.88;         // Interface
         pest.dSlope     = 1.22;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 5.0;
         }
         else
         {
			 pest.dSafetyFactor = 5.0; 
         }

         m_aPestData.Add( pest );

		 
	 // Adds pest Data for 40 degrees 
	 pest.dTemp      = 40;             // Temp
         pest.dNValue    = 1.0;			  // N
         pest.dIntercept = -2.88;         // Interface
         pest.dSlope     = 1.22;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 4.0;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );




	 
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}



//David S 03/13/2007: Enhancement 2005-2007 Added new Pest
void DDasPestData::PopulateCocoaMoth( const LIFE_STAGE &eLifeStage, 
                                           const TREATMENT_TYPE &eTreatmentType,
                                           const FUMIGATION_TYPE &eFumigationType,
                                           const double &dTimeExposure,
                                           CString &strMsg )
{
   ASSERT( strMsg );  

   PEST_DATA pest;

   // The array will be populated with temperatures going from lowest to highest
  
   
   if( eLifeStage == EGG_ALL )// Dosage = "HIGH"
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Inserts data for pest at 15 degrees
         pest.dTemp      = 15;	// temp
         pest.dNValue    = 1.0; // N
         pest.dIntercept = -9.73;  // Intercept
         pest.dSlope     = 1.49;   // Slop

		 // fumigation Type = Space
         if( eFumigationType == SPACE_FUMIGATION ) 
				{
				   pest.dProbit    = PROBIT_85; 
				   pest.dSafetyFactor = 1.0; 
				}
         else // Fumigation Type = Commodity
			 if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dProbit    = PROBIT_95; 
					pest.dSafetyFactor = 0.71; 
				 }
			 else
				 {
					ASSERT( FALSE );
				 }

         // Adds pest data to array
		 m_aPestData.Add( pest );
         
         
         // Inserts data for pest at 20 degrees
         pest.dTemp      = 20;	// temp
         pest.dNValue    = 1.0; // N
         pest.dIntercept = -12.8;  // Intercept
         pest.dSlope     = 1.95;   // Slop

		 // fumigation Type = Space
         if( eFumigationType == SPACE_FUMIGATION ) 
			 {
				pest.dProbit    = PROBIT_85; 
				pest.dSafetyFactor = 1.0; 
			 }
         else // Fumigation Type = Commodity
			 if( eFumigationType == COMMODITY_FUMIGATION )
				 {
					pest.dProbit    = PROBIT_95; 
					pest.dSafetyFactor = 0.8; 
				 }
			 else
				 {
					ASSERT( FALSE );
				 }

         // Adds pest data to array
		 m_aPestData.Add( pest );


	 // Inserts data for pest at 25 degrees
         pest.dTemp      =  25;
         pest.dNValue    =  1.0;
         pest.dIntercept = -12.5;
         pest.dSlope     = 2.0;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 1.0; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


	 // Inserts data for pest at 30 degrees
         pest.dTemp      =  30;
         pest.dNValue    =  1.0;
         pest.dIntercept = -5.3295;
         pest.dSlope     = 1.5141;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 7.95; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 7.5; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );



	 // Inserts data for pest at 35 degrees
         pest.dTemp      =  35;
         pest.dNValue    =  1.0;
         pest.dIntercept = -5.3295;
         pest.dSlope     = 1.5141;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 6.3; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 4.9; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


	 // Inserts data for pest at 40 degrees
         pest.dTemp      =  40;
         pest.dNValue    =  1.0;
         pest.dIntercept = -5.3295;
         pest.dSlope     = 1.5141;
         if( eFumigationType == SPACE_FUMIGATION )
         {
            pest.dProbit    = PROBIT_85; 
            pest.dSafetyFactor = 5.5; 
         }
         else
         if( eFumigationType == COMMODITY_FUMIGATION )
         {
            pest.dProbit    = PROBIT_95; 
            pest.dSafetyFactor = 3.3; 
         }
         else
         {
            ASSERT( FALSE );
         }

		m_aPestData.Add( pest );


 }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   if( eLifeStage == LARVE_PUPAE_ADULT )
   {
      // Pest data for eggs
      if( eTreatmentType == VACUUM || eTreatmentType == NORMAL_ATMOSPHERIC_PRESSURE )
      {
         // Adds pest Data for 15 degrees 
	 pest.dTemp      = 15;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -8.72;         // Interface
         pest.dSlope     = 1.92;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 2.5;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );

         // Adds pest Data for 20 degrees 
	 pest.dTemp      = 20;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -8.2;         // Interface
         pest.dSlope     = 1.83;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );

	
	 // Adds pest Data for 25 degrees 
	 pest.dTemp      = 25;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -7.25;         // Interface
         pest.dSlope     = 1.72;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
		 // Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 2.0;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );	

	 // Adds pest Data for 30 degrees 
	 pest.dTemp      = 30;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -7.25;         // Interface
         pest.dSlope     = 1.72;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
	// Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 1.5;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );

	 // Adds pest Data for 35 degrees 
	 pest.dTemp      = 35;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -7.25;         // Interface
         pest.dSlope     = 1.72;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
	// Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 1.2;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );


	 // Adds pest Data for 40 degrees 
	 pest.dTemp      = 40;         // Temp
         pest.dNValue    = 1.0;        // N
         pest.dIntercept = -7.25;         // Interface
         pest.dSlope     = 1.72;          // Slope
         pest.dProbit    = PROBIT_95;  // Probit
		 
	// Assigns Safty Factor
         if( dTimeExposure >= RANGE_TIME_START && dTimeExposure <= RANGE_TIME_END )
         {
			 pest.dSafetyFactor = 1.0;
         }
         else
         {
			 pest.dSafetyFactor = 4.0; 
         }

         m_aPestData.Add( pest );
	 
      }
      else
      {
         ASSERT( FALSE );
      }
   }
   else
   {
      ASSERT( FALSE );
   }
}































void DDasPestData::GetPestNames(CStringArray &astrPestNames)
{
   // Add all of the pests this calculator supports
   // Ensure the pest table is empty
   astrPestNames.RemoveAll();
   
   // Ensure the order that the pests appear in the array are consistant
   // with the enum PEST.  See PestData.h for the order.

   CString strPest;
   strPest.LoadString( IDS_CONFUSED_FLOUR_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_RED_FLOUR_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_SAWTOOTHED_GRAIN_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_WAREHOUSE_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_INDIAN_MEAL_MOTH );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_MEDITERRANEAN_FLOUR_MOTH );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_CODLING_MOTH );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_NAVEL_ORANGEWORM );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_OTHER_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_OTHER_MOTH );
   astrPestNames.Add( strPest );

   //JNM 04/01/03 - Fum. Enhancements 2003 - Added 2 new pests below.
   strPest.LoadString( IDS_GRANARY_WEEVIL );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_RICE_WEEVIL );
   astrPestNames.Add( strPest );

   //Pat 4/6/05: Added new pest
   strPest.LoadString( IDS_LESSER_GRAIN_BORER );
   astrPestNames.Add( strPest );

   //Pat 12/01/05: Enhancement 05-06 - Added new pests (5)
   strPest.LoadString( IDS_HIDE_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_DRUG_STORE_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_RUSTY_GRAIN_BEETLE );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_ALMOND_MOTH );
   astrPestNames.Add( strPest );

   strPest.LoadString( IDS_RODENTS );
   astrPestNames.Add( strPest );

   //03-13-2007 David Smith Added new pests
   strPest.LoadString( IDS_CIGARETTE_BEETLE );
   astrPestNames.Add( strPest );

   //03-13-2007 David Smith Added new pests
   strPest.LoadString( IDS_COWPEA_WEEVIL );
   astrPestNames.Add( strPest );
     
   //03-13-2007 David Smith Added new pests
   strPest.LoadString( IDS_BEAN_WEEVIL );
   astrPestNames.Add( strPest );
      
   //03-13-2007 David Smith Added new pests
   strPest.LoadString( IDS_COCA_MOTH );
   astrPestNames.Add( strPest );




}
