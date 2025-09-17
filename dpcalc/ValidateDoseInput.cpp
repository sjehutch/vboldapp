// ValidateDoseInput.cpp: implementation of the DValidateDoseInput class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "DasCommodityData.h"
#include "DasFumigationTypeData.h"
#include "DasLifeStageData.h"
#include "DasPestData.h"
#include "DasTreatmentTypeData.h"
#include "GlobalEnums.h"
#include "ValidateDoseInput.h"


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif



//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DValidateDoseInput::DValidateDoseInput()
{

}


DValidateDoseInput::~DValidateDoseInput()
{

}


BOOL DValidateDoseInput::ValidateDoseInput( LPDOSE_INPUT_DATA lpdataDoseInput )
{
   ASSERT( lpdataDoseInput );

   // Time to validate the user input data to ensure it is valid prior to performing the calcs
   if( !ValidatePest( lpdataDoseInput->nPest, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateFumigationType( lpdataDoseInput->nFumigationType, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateLifeStage( lpdataDoseInput->nLifeStage, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateTreatmentType( lpdataDoseInput->nTreatmentType, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateCommodity( lpdataDoseInput->nCommodity, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateExposureTime( lpdataDoseInput->dTimeExposure, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateHlt( lpdataDoseInput->dHlt, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateTemp( lpdataDoseInput->dTemp, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   if( !ValidateVolume( lpdataDoseInput->dVolume, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

 
   if( !ValidateLoadFactor( lpdataDoseInput->dLoadFactor, lpdataDoseInput->strError ) )
   {
      return FALSE;
   }

   return TRUE;
}


BOOL DValidateDoseInput::ValidatePest( const int &nPest, CString &strError )
{
   BOOL bReturn = TRUE;

   switch( nPest )
   {
   case PEST_CONFUSED_FLOUR_BEETLE:
   case PEST_RED_FLOUR_BEETLE:
   case PEST_SAWTOOTHED_GRAIN_BEETLE:
   case PEST_WAREHOUSE_BEETLE:
   case PEST_INDIAN_MEAL_MOTH:
   case PEST_MEDITERRANEAN_FLOUR_MOTH:
   case PEST_CODLING_MOTH:
   case PEST_NAVEL_ORANGEWORM:
   case PEST_OTHER_BEETLE:
   case PEST_OTHER_MOTH:
   //JNM 04/01/03 - Fum. Enhancements 2003 - Added 2 new cases for 2 new pests.
   case PEST_GRANARY_WEEVIL:
   case PEST_RICE_WEEVIL:
   //Pat 4/6/05: Added new pest
   case PEST_LESSER_GRAIN_BORER:
   //Pat 12/2/05 Added new pests (5)
   case PEST_HIDE_BEETLE: 
   case PEST_DRUG_STORE_BEETLE: 
   case PEST_RUSTY_GRAIN_BEETLE:
   case PEST_ALMOND_MOTH:
   case PEST_RODENTS:
   //DavidSmith 03/13/07 Added new pests
   case PEST_CIGARETTE_BEETLE:
   case PEST_COWPEA_WEEVIL:
   case PEST_BEAN_WEEVIL:
   case PEST_COCOA_MOTH:



   break;

   default:
      strError.Format( IDS_INVALID_PEST, PEST_CONFUSED_FLOUR_BEETLE, PEST_RICE_WEEVIL );
      bReturn = FALSE;
   }

   return bReturn;;
}


BOOL DValidateDoseInput::ValidateCommodity( const int &nCommodity, CString &strError )
{
   BOOL bReturn = TRUE;

   switch( nCommodity )
   {
   case COMMODITY_NONE:
   case COMMODITY_WALNUTS:
   case COMMODITY_ALMONDS:
   case COMMODITY_FIGS:
   // Pat 1/16/06 - Removed COMMODITY_FILBERTS_HAZELNUTS changed to APRICOTS_DRIED
   case COMMODITY_APRICOTS_DRIED:
   case COMMODITY_PISTACHIOS:
   // Pat 1/16/06 - Removed COMMODITY_PRUNES changed to CACAO
   case COMMODITY_CACAO:
   case COMMODITY_RAISINS:
   // Pat 1/16/06 - Removed COMMODITY_WHOLE_GRAIN changed to COFFEE
   case COMMODITY_COFFEE:
   //JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new case for Other commmodity.
   // Pat 1/16/06 - Removed COMMODITY_OTHER changed to CORN
   //case COMMODITY_OTHER:
   case COMMODITY_CORN:
   case COMMODITY_COWPEA:
   case COMMODITY_DATES:
   case COMMODITY_OTHER_DRIED_FRUIT:
   case COMMODITY_OTHER_GRAIN:
   case COMMODITY_OTHER_LEGUME:
   case COMMODITY_OTHER_NUTS:
   case COMMODITY_PECANS:
   case COMMODITY_PLUMS_DRIED:
   case COMMODITY_POPCORN:
   case COMMODITY_RICE_POLISHED:
   case COMMODITY_RICE_ROUGH:
   case COMMODITY_SOYBEAN:
   case COMMODITY_WHEAT:
      break;

   default:
      strError.Format( IDS_INVALID_COMMODITY, COMMODITY_NONE );
      bReturn = FALSE;
   }

   return bReturn;;
}


BOOL DValidateDoseInput::ValidateFumigationType( const int &nFumigationType, CString &strError )
{
   BOOL bReturn = TRUE;

   switch( nFumigationType )
   {
   case SPACE_FUMIGATION:
   case COMMODITY_FUMIGATION:
      break;

   default:
      strError.Format( IDS_INVALID_FUMIGATION_TYPE, SPACE_FUMIGATION, COMMODITY_FUMIGATION );
      bReturn = FALSE;
   }

   return bReturn;
}


BOOL DValidateDoseInput::ValidateLifeStage( const int &nLifeStage, CString &strError )
{
   BOOL bReturn = TRUE;

   switch( nLifeStage )
   {
   case EGG_ALL:
   case LARVE_PUPAE_ADULT:
      break;

   default:
      strError.Format( IDS_INVALID_LIFE_STAGE, EGG_ALL, LARVE_PUPAE_ADULT );
      bReturn = FALSE;
   }

   return bReturn;
}


BOOL DValidateDoseInput::ValidateTreatmentType( const int &nTreatmentType, CString &strError )
{
   BOOL bReturn = TRUE;

   switch( nTreatmentType )
   {
   case NORMAL_ATMOSPHERIC_PRESSURE:
   case VACUUM:
      break;

   default:
      strError.Format( IDS_INVALID_TREATMENT_TYPE, NORMAL_ATMOSPHERIC_PRESSURE, VACUUM );
      bReturn = FALSE;
   }

   return bReturn;
}


BOOL DValidateDoseInput::ValidateHlt( const double &dHlt, CString &strError )
{
   BOOL bReturn = TRUE;

   if( dHlt < MIN_HLT )
   {
      strError.Format( IDS_INVALID_HLT, MIN_HLT );
      bReturn = FALSE;
   }

   return bReturn;
}


BOOL DValidateDoseInput::ValidateTemp( const double &dTemp, CString &strError )
{
   BOOL bReturn = TRUE;

   if( dTemp < MIN_TEMP || dTemp > MAX_TEMP )
   {
      strError.Format( IDS_INVALID_TEMP, MIN_TEMP, MAX_TEMP );
      bReturn = FALSE;
   }

   return bReturn;
}


BOOL DValidateDoseInput::ValidateExposureTime( const double &dTimeExposure, CString &strError )
{
   BOOL bReturn = TRUE;

   if( dTimeExposure < MIN_TIME )
   {
      strError.Format( IDS_INVALID_TIME, MIN_TIME );
      bReturn = FALSE;
   }

   return bReturn;
}


BOOL DValidateDoseInput::ValidateVolume( const double &dVolume, CString &strError )
{
   BOOL bReturn = TRUE;

   if( dVolume < MIN_VOL )
   {
      strError.Format( IDS_INVALID_VOLUME, MIN_VOL );
      bReturn = FALSE;
   }

   return bReturn;
}


BOOL DValidateDoseInput::ValidateLoadFactor( const double &dLoadFactor, CString &strError )
{
   BOOL bReturn = TRUE;

   if( dLoadFactor < MIN_LOAD_FACTOR )
   {
      strError.Format( IDS_INVALID_LOAD_FACTOR, MIN_LOAD_FACTOR );
      bReturn = FALSE;
   }

   return bReturn;
}

