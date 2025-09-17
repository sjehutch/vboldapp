// ValidateAreaMonPtInput.cpp: implementation of the DValidateAreaMonPtInput class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "ValidateAreaMonPtInput.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DValidateAreaMonPtInput::DValidateAreaMonPtInput()
{

}

DValidateAreaMonPtInput::~DValidateAreaMonPtInput()
{

}


BOOL DValidateAreaMonPtInput::ValidateAreaMonPt( LPAREA_MON_PT_DATA lpAreaMonPtData )
{
   ASSERT( lpAreaMonPtData );

   BOOL bReturn = TRUE;

   // Validate all parts of the structure
   if( !ValidateAchievedConcentrationTime( lpAreaMonPtData->dAchievedConcentrationTime,
                                           lpAreaMonPtData->strMsg ) )
   {
      return FALSE;
   }

   if( !ValidateAchievedDose( lpAreaMonPtData->dAchievedDose,
                              lpAreaMonPtData->strMsg ) )
   {
      return FALSE;
   }

   
   return bReturn;
}


BOOL DValidateAreaMonPtInput::ValidateAchievedConcentrationTime( double &dConcentrationTime, CString &strError )
{
   BOOL bReturn = TRUE;

   if( dConcentrationTime < 0.0 )
   {
      strError.LoadString( IDS_INVALID_CONC_TIME );

      return FALSE;
   }
   
   return bReturn;
}


BOOL DValidateAreaMonPtInput::ValidateAchievedDose( double &dDose, CString &strError )
{
   BOOL bReturn = TRUE;
   
   if( dDose < 0.0 )
   {
      strError.LoadString( IDS_INVALID_DOSE );

      return FALSE;
   }
   
   return bReturn;
}
