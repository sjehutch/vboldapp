// DasFumigationTypeData.cpp: implementation of the DDasFumigationTypeData class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DasFumigationTypeData.h"
#include "resource.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DDasFumigationTypeData::DDasFumigationTypeData()
{

}

DDasFumigationTypeData::~DDasFumigationTypeData()
{

}

void DDasFumigationTypeData::GetFumigationTypeNames(CStringArray &astrFumigationTypeNames)
{
   // Add all of the fumigation types this calculator supports
   // Ensure the fumigation type table is empty
   astrFumigationTypeNames.RemoveAll();
   
   // Ensure the order that the fumigation types appear in the array are consistant
   // with the enum FUMIGATION_TYPE.  See DasFumigationTypeData.h for the order.

   CString strFumigationType;
   strFumigationType.LoadString( IDS_SPACE );
   astrFumigationTypeNames.Add( strFumigationType );

   strFumigationType.LoadString( IDS_COMMODITY );
   astrFumigationTypeNames.Add( strFumigationType );
}
