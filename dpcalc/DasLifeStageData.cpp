// DasLifeStageData.cpp: implementation of the DDasLifeStageData class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DasLifeStageData.h"
#include "resource.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DDasLifeStageData::DDasLifeStageData()
{

}

DDasLifeStageData::~DDasLifeStageData()
{

}

void DDasLifeStageData::GetLifeStageNames(CStringArray &astrLifeStageNames)
{
   // Add all of the fumigation types this calculator supports
   // Ensure the fumigation type table is empty
   astrLifeStageNames.RemoveAll();
   
   // Ensure the order that the fumigation types appear in the array are consistant
   // with the enum FUMIGATION_TYPE.  See DasFumigationTypeData.h for the order.

   CString strLifeStage;
   strLifeStage.LoadString( IDS_LARVE_PUPAE_ADULT );
   astrLifeStageNames.Add( strLifeStage );

   strLifeStage.LoadString( IDS_EGG_ALL );
   astrLifeStageNames.Add( strLifeStage );
}
