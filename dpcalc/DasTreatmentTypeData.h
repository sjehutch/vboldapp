// DasTreatmentTypeData.h: interface for the DDasTreatmentTypeData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DASTREATMENTTYPEDATA_H__2F2FEBC5_FEE1_40E2_95DD_A54F0D8F9088__INCLUDED_)
#define AFX_DASTREATMENTTYPEDATA_H__2F2FEBC5_FEE1_40E2_95DD_A54F0D8F9088__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


// =DOC SECTION enum TREATMENT_TYPE
// =DOC SECTION Pest treatment type enumeration
// =DOC TEXT enum TREATMENT_TYPE extends nothing
// =DOC TEXT Pest treatment type enumeration
// =DOC TEXT NORMAL_ATMOSPHERIC_PRESSURE = 0
// =DOC TEXT VACUUM                      = 1
//////////////////////////////////////////////////////////////////
enum TREATMENT_TYPE { NORMAL_ATMOSPHERIC_PRESSURE = 0,
                      VACUUM                      = 1 };


// =DOC SECTION Class DDasTreatmentTypeData
// =DOC SECTION Treatment type data handler
// =DOC TEXT class DDasTreatmentTypeData extends nothing.
// =DOC TEXT The class contains the treatment type data used in all algorithms.
//////////////////////////////////////////////////////////////////
class DDasTreatmentTypeData  
{
public:
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DDasTreatmentTypeData();


   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
	virtual ~DDasTreatmentTypeData();

   // =DOC SECTION Method GetTreatmentTypeNames( CStringArray &astrTreatmentTypeNames ): void
	// =DOC SECTION Gets the treatment type names and places them in the astrTreatmentTypeNames parameter 
   // =DOC TEXT public  GetTreatmentTypeNames( CStringArray &astrTreatmentTypeNames ): void
   // =DOC TEXT CStringArray &astrCommodityNames: Array of commodity names
	// =DOC TEXT Returns void: No return value
   // =DOC TEXT The functions retrieves a list of treatment type names in which data exists
   //////////////////////////////////////////////////////////////////
	void GetTreatmentTypeNames( CStringArray &astrTreatmentTypeNames );
};

#endif // !defined(AFX_DASTREATMENTTYPEDATA_H__2F2FEBC5_FEE1_40E2_95DD_A54F0D8F9088__INCLUDED_)
