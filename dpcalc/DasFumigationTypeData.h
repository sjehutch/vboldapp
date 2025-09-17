// DasFumigationTypeData.h: interface for the DDasFumigationTypeData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DASFUMIGATIONTYPEDATA_H__E0A4A884_25B0_460A_8585_B717B1A85476__INCLUDED_)
#define AFX_DASFUMIGATIONTYPEDATA_H__E0A4A884_25B0_460A_8585_B717B1A85476__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



// =DOC SECTION enum FUMIGATION_TYPE
// =DOC SECTION Fumigation type enumeration
// =DOC TEXT enum FUMIGATION_TYPE extends nothing
// =DOC TEXT Fumigation type enumeration
// =DOC TEXT SPACE_FUMIGATION     = 0
// =DOC TEXT COMMODITY_FUMIGATION = 1
//////////////////////////////////////////////////////////////////
enum FUMIGATION_TYPE { SPACE_FUMIGATION     = 0,
                       COMMODITY_FUMIGATION = 1 };


// =DOC SECTION Class DDasFumigationTypeData
// =DOC SECTION Fumigation type data handler
// =DOC TEXT class DDasFumigationTypeData extends nothing.
// =DOC TEXT The class contains the fumigation data used in all algorithms.
//////////////////////////////////////////////////////////////////
class DDasFumigationTypeData  
{
public:
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DDasFumigationTypeData();

   
   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DDasFumigationTypeData();

   // =DOC SECTION Method GetFumigationTypeNames( CStringArray &astrFumigationTypeNames ): void
	// =DOC SECTION Gets the fumigation type names and places them in the astrFumigationTypeNames parameter 
   // =DOC TEXT public  GetFumigationTypeNames( CStringArray &astrFumigationTypeNames ): void
   // =DOC TEXT CStringArray &astrFumigationTypeNames: Array of fumigation type names
	// =DOC TEXT Returns void: No return value
   // =DOC TEXT The functions retrieves a list of fumigation type names in which data exists
   //////////////////////////////////////////////////////////////////
	void GetFumigationTypeNames( CStringArray &astrFumigationTypeNames );
};

#endif // !defined(AFX_DASFUMIGATIONTYPEDATA_H__E0A4A884_25B0_460A_8585_B717B1A85476__INCLUDED_)
