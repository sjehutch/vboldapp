// DasCommodityData.h: interface for the DCommodityData class.
// JNM 04/09/03 - Fumiguide Enh. 2003 - Added new commodity.
//									  - Added function declarations for the
//									    new commodity.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DASCOMMODITYDATA_H__B25A7A65_EE6B_42A7_973E_4D0203BC7FDE__INCLUDED_)
#define AFX_DASCOMMODITYDATA_H__B25A7A65_EE6B_42A7_973E_4D0203BC7FDE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "CommodityData.h"	// Added by ClassView
#include "DasFumigationTypeData.h"
#include "DasTreatmentTypeData.h"
#include "PestData.h"


// =DOC SECTION enum COMMODITY
// =DOC SECTION Commodity enumeration
// =DOC TEXT enum COMMODITY extends nothing
// =DOC TEXT Commodity enumeration
// =DOC TEXT COMMODITY_NONE               = 0
// =DOC TEXT COMMODITY_ALMONDS            = 1
// =DOC TEXT COMMODITY_FIGS               = 2
// Pat 1/16/06 - Removed COMMODITY_FILBERTS_HAZELNUTS changed to APRICOTS_DRIED
// =DOC TEXT COMMODITY_APRICOTS_DRIED	  = 3
// =DOC TEXT COMMODITY_PISTACHIOS         = 4
// Pat 1/16/06 - Removed COMMODITY_PRUNES changed to CACAO
// =DOC TEXT COMMODITY_CACAO              = 5
// =DOC TEXT COMMODITY_RAISINS            = 6
// =DOC TEXT COMMODITY_WALNUTS            = 7
// Pat 1/16/06 - Removed COMMODITY_WHOLE_GRAIN changed to COFFEE
// =DOC TEXT COMMODITY_COFFEE             = 8
// JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new commodity below.
// Pat 1/16/06 - Removed COMMODITY_OTHER changed to CORN
// =DOC TEXT COMMODITY_CORN  			  = 9
// Pat 1/16/06 - Added the following COMMODITIES
// =DOC TEXT COMMODITY_COWPEA  			  = 10
// =DOC TEXT COMMODITY_DATES  			  = 11
// =DOC TEXT COMMODITY_OTHER_DRIED_FRUIT  = 12
// =DOC TEXT COMMODITY_OTHER_GRAIN        = 13
// =DOC TEXT COMMODITY_OTHER_LEGUME       = 14
// =DOC TEXT COMMODITY_OTHER_NUTS         = 15
// =DOC TEXT COMMODITY_PECANS             = 16
// =DOC TEXT COMMODITY_PLUMS_DRIED        = 17
// =DOC TEXT COMMODITY_POPCORN            = 18
// =DOC TEXT COMMODITY_RICE_POLISHED      = 19
// =DOC TEXT COMMODITY_RICE_ROUGH         = 20
// =DOC TEXT COMMODITY_SOYBEAN            = 21
// =DOC TEXT COMMODITY_WHEAT              = 22
//PAT 1/16/06 -  ENd of NEW Commodities
//////////////////////////////////////////////////////////////////
enum COMMODITY { COMMODITY_NONE               = 0,
                 COMMODITY_ALMONDS            = 1,
                 COMMODITY_FIGS               = 2,
                 COMMODITY_APRICOTS_DRIED     = 3,
                 COMMODITY_PISTACHIOS         = 4,
                 COMMODITY_CACAO              = 5,
                 COMMODITY_RAISINS            = 6,
                 COMMODITY_WALNUTS            = 7,
                 COMMODITY_COFFEE             = 8,
				 COMMODITY_CORN  			  = 9,
				 COMMODITY_COWPEA  			  = 10,
				 COMMODITY_DATES  			  = 11,
				 COMMODITY_OTHER_DRIED_FRUIT  = 12,
				 COMMODITY_OTHER_GRAIN        = 13,
				 COMMODITY_OTHER_LEGUME       = 14,
				 COMMODITY_OTHER_NUTS         = 15,
				 COMMODITY_PECANS             = 16,
				 COMMODITY_PLUMS_DRIED        = 17,
				 COMMODITY_POPCORN            = 18,
				 COMMODITY_RICE_POLISHED      = 19,
				 COMMODITY_RICE_ROUGH         = 20,
				 COMMODITY_SOYBEAN            = 21,
				 COMMODITY_WHEAT              = 22 };


// =DOC SECTION Class DDasCommodityData
// =DOC SECTION Commodity data handler
// =DOC TEXT class DDasCommodityData extends nothing.
// =DOC TEXT The class contains the commodity data used in all algorithms.
//////////////////////////////////////////////////////////////////
class DDasCommodityData  
{
public:
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DDasCommodityData();


   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
	virtual ~DDasCommodityData();


   // =DOC SECTION Method GetCommodityNames( CStringArray &astrCommodityNames ): void
	// =DOC SECTION Gets the commodity names and places them in the astrCommodityNames parameter 
   // =DOC TEXT public  GetCommodityNames( CStringArray &astrCommodityNames ): void
   // =DOC TEXT CStringArray &astrCommodityNames: Array of commodity names
	// =DOC TEXT Returns void: No return value
   // =DOC TEXT The functions retrieves a list of commodity names in which data exists
   //////////////////////////////////////////////////////////////////
	void GetCommodityNames( CStringArray &astrCommodityNames );


   // =DOC SECTION Method PopulateCommodityData( const COMMODITY &eCommodity, const FUMIGATION_TYPE &eFumigationType, const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populates the internal member m_dataCommodity for a particular commodity 
   // =DOC TEXT public  PopulateCommodityData( const COMMODITY &eCommodity, const FUMIGATION_TYPE &eFumigationType, const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
   // =DOC TEXT COMMODITY &eCommodity: The commodity type
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
   // =DOC TEXT The functions retrieves the commodity data for a requested commodity
   //////////////////////////////////////////////////////////////////
	void PopulateCommodityData( const COMMODITY &eCommodity, 
                               const FUMIGATION_TYPE &eFumigationType, 
                               const TREATMENT_TYPE &eTreatmentType, 
                               CString &strMsg );
	
   // =DOC SECTION Method GetCommodityData(): COMMODITY_DATA
	// =DOC SECTION Retrieves the commodity data including maximum permitted CT for a requested commodity
   // =DOC TEXT public  GetCommodityData(): COMMODITY_DATA
   // =DOC TEXT Returns COMMODITY_DATA: Structure containing commodity data including maximum permitted CT for a particular commodity
   // =DOC TEXT The functions retrieves the commodity data for a requested commodity
   //////////////////////////////////////////////////////////////////
   COMMODITY_DATA GetCommodityData();

protected:
	// =DOC SECTION Method PopulateNone( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with none data.  This is also selected if it is a space fumigation.
	// =DOC TEXT protected  PopulateWholeGrain( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with none data.  This is also selected if it is a space fumigation.
	//////////////////////////////////////////////////////////////////
	void PopulateNone( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

    // =DOC SECTION Method PopulateAlmonds( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with almond data
	// =DOC TEXT protected  PopulateAlmonds( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with almond data
	//////////////////////////////////////////////////////////////////
	void PopulateAlmonds( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );
	
    // =DOC SECTION Method PopulateFigs( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with fig data
	// =DOC TEXT protected  PopulateFigs( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with fig data
	//////////////////////////////////////////////////////////////////
	void PopulateFigs( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );
	
	//Pat 1/17/2006 - Removed PopulateFilbertsHazelnuts function

    // =DOC SECTION Method PopulateFilbertsHazelnuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with filbert/hazelnut data
	// =DOC TEXT protected  PopulateFilbertsHazelnuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with filbert/hazelnut data
	//////////////////////////////////////////////////////////////////
	//void PopulateFilbertsHazelnuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateApricotsDried function

    // =DOC SECTION Method PopulateApricotsDried( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with apricots data
	// =DOC TEXT protected  PopulateApricotsDried( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with filbert/hazelnut data
	//////////////////////////////////////////////////////////////////
	void PopulateApricotsDried( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );
	
    // =DOC SECTION Method PopulatePistachios( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with pistachio data
	// =DOC TEXT protected  PopulatePistachios( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with pistachio data
	//////////////////////////////////////////////////////////////////
	void PopulatePistachios( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );
	
    //Pat 1/17/2006 - Removed PopulatePrunes function
	
	// =DOC SECTION Method PopulatePrunes( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with prune data
	// =DOC TEXT protected  PopulatePrunes( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	//void PopulatePrunes( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );
	
    //Pat 1/17/2006 - Added PopulateCacao function
	
	// =DOC SECTION Method PopulateCacao( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with cacao data
	// =DOC TEXT protected  PopulateCacao( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateCacao( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

    // =DOC SECTION Method PopulateRaisins( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with raisin data
	// =DOC TEXT protected  PopulateRaisins( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with raisin data
	//////////////////////////////////////////////////////////////////
	void PopulateRaisins( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

    // =DOC SECTION Method PopulateWalnuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with walnut data
	// =DOC TEXT protected  PopulateWalnuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with walnut data
	//////////////////////////////////////////////////////////////////
	void PopulateWalnuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );
	
	//Pat 1/17/2006 - Removed PopulateWholeGrain function

	// =DOC SECTION Method PopulateWholeGrain( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with whole grain data
	// =DOC TEXT protected  PopulateWholeGrain( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with whole grain data
	//////////////////////////////////////////////////////////////////
	//void PopulateWholeGrain( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

    //Pat 1/17/2006 - Added PopulateCoffee function
	
	// =DOC SECTION Method PopulateCoffee( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with coffee data
	// =DOC TEXT protected  PopulateCoffee( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateCoffee( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Removed PopulateOther function

	//JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new method for new commodity.
	// =DOC SECTION Method PopulateOther( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with other data
	// =DOC TEXT protected  PopulateOther( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with other data
	//////////////////////////////////////////////////////////////////
	//void PopulateOther( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateCorn function
	
	// =DOC SECTION Method PopulateCorn( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with corn data
	// =DOC TEXT protected  PopulateCorn( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateCorn( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateCowpea function
	
	// =DOC SECTION Method PopulateCowpea( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with cowpea data
	// =DOC TEXT protected  PopulateCowpea( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateCowpea( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateDates function
	
	// =DOC SECTION Method PopulateDates( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with dates data
	// =DOC TEXT protected  PopulateDates( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateDates( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );
	
	//Pat 1/17/2006 - Added PopulateOtherDriedFruit function
	
	// =DOC SECTION Method PopulateOtherDriedFruit( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with other dried fruits data
	// =DOC TEXT protected  PopulateOtherDriedFruit( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateOtherDriedFruit( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateOtherGrain function
	
	// =DOC SECTION Method PopulateOtherGrain( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with other grains data
	// =DOC TEXT protected  PopulateOtherGrain( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateOtherGrain( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateOtherLegume function
	
	// =DOC SECTION Method PopulateOtherLegume( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with other legumes data
	// =DOC TEXT protected  PopulateOtherLegume( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateOtherLegume( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateOtherNuts function
	
	// =DOC SECTION Method PopulateOtherNuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with other nuts data
	// =DOC TEXT protected  PopulateOtherNuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateOtherNuts( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateOtherPecans function
	
	// =DOC SECTION Method PopulatePecans( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with other pecans data
	// =DOC TEXT protected  PopulatePecans( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulatePecans( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulatePlumsDried function
	
	// =DOC SECTION Method PopulatePlumsDried( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with plums (dried) data
	// =DOC TEXT protected  PopulatePlumsDried( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulatePlumsDried( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulatePopcorn function
	
	// =DOC SECTION Method PopulatePopcorn( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with popcorn data
	// =DOC TEXT protected  PopulatePopcorn( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulatePopcorn( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateRicePolished function
	
	// =DOC SECTION Method PopulateRicePolished( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with rice (polished) data
	// =DOC TEXT protected  PopulateRicePolished( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateRicePolished( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateRiceRough function
	
	// =DOC SECTION Method PopulateRiceRough( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with rice (rough) data
	// =DOC TEXT protected  PopulateRiceRough( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateRiceRough( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateSoybean function
	
	// =DOC SECTION Method PopulateSoybean( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with soybean data
	// =DOC TEXT protected  PopulateSoybean( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateSoybean( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );

	//Pat 1/17/2006 - Added PopulateWheat function
	
	// =DOC SECTION Method PopulateWheat( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_dataCommodity with wheat data
	// =DOC TEXT protected  PopulateWheat( const TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
    // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the fumigation
    // =DOC TEXT CString &strMsg: Error message if problems occur
 	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate commodity data (m_dataCommodity) with prune data
	//////////////////////////////////////////////////////////////////
	void PopulateWheat( const TREATMENT_TYPE &eTreatmentType, CString &strMsg );



	// =DOC SECTION Member m_dataCommodity: COMMODITY_DATA
	// =DOC SECTION Commodity data
	// =DOC TEXT m_dataCommodity: COMMODITY_DATA
	// =DOC TEXT Commodity data populated by the Populate* functions
	//////////////////////////////////////////////////////////////////
	COMMODITY_DATA m_dataCommodity;
};

#endif // !defined(AFX_DASCOMMODITYDATA_H__B25A7A65_EE6B_42A7_973E_4D0203BC7FDE__INCLUDED_)
