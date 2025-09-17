// ValidateConcentration.cpp: implementation of the DValidateConcentration class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "ValidateConcentration.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DValidateConcentration::DValidateConcentration()
{

}


DValidateConcentration::~DValidateConcentration()
{

}


BOOL DValidateConcentration::ValidateConcentration(LPCONC_INPUT_DATA lpConcInputData)
{
   if( !ValidateConcentrationValue( lpConcInputData->dConcentration, lpConcInputData->strMsg ) )
   {
      return FALSE;
   }

   if( !ValidateDate( lpConcInputData->odtDateTime, lpConcInputData->strMsg ) )
   {
      return FALSE;
   }

   return TRUE;
}


BOOL DValidateConcentration::ValidateConcentrationValue( const double &dConcentration, CString &strMsg )
{
   if( dConcentration < 0.0 )
   {
      strMsg = _T( "Error: Concentration must be greater than 0.0" );

      return FALSE;
   }

   return TRUE;
}



BOOL DValidateConcentration::ValidateDate( const COleDateTime &odtDateTime, CString &strMsg )
{
   if( odtDateTime.m_status == COleDateTime::invalid )
   {
      strMsg = _T( "Error: Invalid date/time for the concentration." );
      
      return FALSE;
   }

   return TRUE;
}
