// DasTreatmentTypeData.cpp: implementation of the DDasTreatmentTypeData class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DasTreatmentTypeData.h"
#include "resource.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

DDasTreatmentTypeData::DDasTreatmentTypeData()
{

}

DDasTreatmentTypeData::~DDasTreatmentTypeData()
{

}

void DDasTreatmentTypeData::GetTreatmentTypeNames(CStringArray &astrTreatmentTypeNames)
{
   // Add all of the treatment types this calculator supports
   // Ensure the treatment type table is empty
   astrTreatmentTypeNames.RemoveAll();
   
   // Ensure the order that the treatment types appear in the array are consistant
   // with the enum TREATMENT_TYPE.  See DasTreatmentTypeData.h for the order.

   CString strTreatmentType;
   strTreatmentType.LoadString( IDS_NORMAL_ATMOSPHERIC_PRESSURE );
   astrTreatmentTypeNames.Add( strTreatmentType );

   strTreatmentType.LoadString( IDS_VACUUM );
   astrTreatmentTypeNames.Add( strTreatmentType );
}
