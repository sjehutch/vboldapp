// DasLifeStageData.h: interface for the DDasLifeStageData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DASLIFESTAGEDATA_H__F9D45614_6E9B_4727_A581_CFF4CA3C61B4__INCLUDED_)
#define AFX_DASLIFESTAGEDATA_H__F9D45614_6E9B_4727_A581_CFF4CA3C61B4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



// =DOC SECTION enum LIFE_STAGE
// =DOC SECTION Pest life stage enumeration
// =DOC TEXT enum LIFE_STAGE extends nothing
// =DOC TEXT Pest life stage enumeration
// =DOC TEXT LARVE_PUPAE_ADULT = 0
// =DOC TEXT EGG_ALL           = 1
//////////////////////////////////////////////////////////////////
enum LIFE_STAGE { LARVE_PUPAE_ADULT = 0,
                  EGG_ALL           = 1 };


// =DOC SECTION Class DDasLifeStageData
// =DOC SECTION Life stage data handler
// =DOC TEXT class DDasLifeStageData extends nothing.
// =DOC TEXT The class contains the life stage data used in all algorithms.
//////////////////////////////////////////////////////////////////
class DDasLifeStageData  
{
public:
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DDasLifeStageData();

   
   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DDasLifeStageData();

   // =DOC SECTION Method GetLifeStageNames( CStringArray &astrLifeStageNames ): void
	// =DOC SECTION Gets the life stage names and places them in the astrLifeStageNames parameter 
   // =DOC TEXT public  GetLifeStageNames( CStringArray &astrLifeStageNames ): void
   // =DOC TEXT CStringArray &astrLifeStageNames: Array of life stage names
	// =DOC TEXT Returns void: No return value
   // =DOC TEXT The functions retrieves a list of life stage names in which data exists
   //////////////////////////////////////////////////////////////////
	void GetLifeStageNames( CStringArray &astrLifeStageNames );
};

#endif // !defined(AFX_DASLIFESTAGEDATA_H__F9D45614_6E9B_4727_A581_CFF4CA3C61B4__INCLUDED_)
