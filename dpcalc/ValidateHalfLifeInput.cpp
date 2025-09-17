// ValidateHalfLifeInput.cpp: implementation of the DValidateHalfLifeInput class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "ValidateHalfLifeInput.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DValidateHalfLifeInput::DValidateHalfLifeInput()
{

}


DValidateHalfLifeInput::~DValidateHalfLifeInput()
{

}


BOOL DValidateHalfLifeInput::ValidateHalfLifeInput( LPHALFLIFE_INPUT_DATA lpdataHalfLifeInput )
{
   ASSERT( lpdataHalfLifeInput );

   if( !ValidateHlt( lpdataHalfLifeInput->dHlt, lpdataHalfLifeInput->strMsg ) )
   {
      return FALSE;
   }

   return TRUE;
}


BOOL DValidateHalfLifeInput::ValidateHlt( double& dHlt, CString& strMsg )
{
   if( dHlt < 0.0 )
   {
      strMsg = _T( "Half-life time must be greater than or equal to 0.0." );

      return FALSE;
   }

   return TRUE;
}
